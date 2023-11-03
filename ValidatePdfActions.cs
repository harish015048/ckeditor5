//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Web;
//using iTextSharp.text.pdf;
//using System.Text;
//using iTextSharp.text.pdf.parser;
//using CMCai.Models;
//using System.IO;
//using iTextSharp.text;
//using System.Text.RegularExpressions;
//using System.Net;
//using iTextSharp.text.pdf.security;
//using System.Security.Cryptography.X509Certificates;
//using Bytescout.PDFExtractor;

//namespace CMCai.Actions
//{
//    public class ValidatePdfActions
//    {
//        string sourcePath1 = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
//        string destPath1 = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
//        string sourcePathFolder = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCDestination/");
//        RegOpsQCActions qObj = new RegOpsQCActions();

//        string sourcePath = string.Empty;
//        string destPath10 = string.Empty;
//        public string Check_PDFVersion(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                //sourcePath = path;
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;                
//                rObj.CHECK_START_TIME = DateTime.Now;
//                PdfReader reader = new PdfReader(sourcePath);
//                StringWriter output = new StringWriter();
//                string version = reader.PdfVersion.ToString();
//                if (version == "4" || version == "5" || version == "6" || version == "7")
//                {
//                    rObj.QC_Result = "Passed";
//                }
//                else
//                    rObj.QC_Result = "Failed";

//                rObj.Comments = "PDF Version is 1." + version + "";
//                output.Close();
//                reader.Close();                
//                res = qObj.SaveValidateResults(rObj);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public void Check_PDFFile_PasswordProtection(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                PdfReader reader = new PdfReader(sourcePath);
//                bool result = reader.IsOpenedWithFullPermissions;
//                reader.Close();
//                if (result == true)
//                {
//                    rObj.QC_Result = "Passed";
//                    rObj.Comments = "There is no password has been set for this file.";
//                }
//                else
//                {
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = "Password has been set for this file.";
//                }
//                rObj.CHECK_END_TIME = DateTime.Now;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                rObj.Job_Status = "Failed";
//                rObj.QC_Result = "Error";
//                rObj.Comments = "Technical error: " + ex.Message;                
//            }
//        }

//        public string PDFFile_Properties(RegOpsQC rObj, string path, string checkString, string destPath, double checkType)
//        {
//            string res = string.Empty;
//            try
//            {
//                // sourcePath = destPath1 + rObj.Job_ID + "/Source/" + rObj.File_Name;
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                PdfReader reader = new PdfReader(sourcePath);
//                IDictionary<String, String> metadic = reader.Info;
//                var dic = from m in metadic
//                          select m;
//                string result = string.Empty;
//                rObj.CHECK_START_TIME = DateTime.Now;

//                if (checkString == "Properties fields should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        //Console.WriteLine(d.Key + ": " + d.Value);
//                        //result = result + " , " + d.Key + ": " + d.Value;
//                        if ((d.Key == "Title" || d.Key == "Subject" || d.Key == "Author" || d.Key == "Keywords") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (checkType == 1)
//                    {
//                        if (result != "")
//                        {
//                            PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                            SortedDictionary<String, String> inf = new SortedDictionary<String, String>();
//                            inf.Add("Title", "");
//                            inf.Add("Subject", "");
//                            inf.Add("Author", "");
//                            inf.Add("Keywords", "");
//                            stamper.MoreInfo = inf;
//                            stamper.Close();
//                            rObj.QC_Result = "Fixed";
//                            rObj.Comments = "Default properties has been set to blank.";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Default properties are empty.";
//                        }
//                    }
//                    else
//                    {
//                        if (result != "")
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Default properties are not empty.";

//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Default properties are empty.";
//                        }
//                    }

//                }
//                else if (checkString == "Title should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Title") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (checkType == 1)
//                    {
//                        if (result != "")
//                        {
//                            PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                            SortedDictionary<String, String> inf = new SortedDictionary<String, String>();
//                            inf.Add("Title", "");
//                            stamper.MoreInfo = inf;
//                            stamper.Close();
//                            rObj.QC_Result = "Fixed";
//                            rObj.Comments = "Title property value is not empty.";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Title property value is empty.";
//                        }
//                    }
//                    else
//                    {
//                        if (result != "")
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Title property value is not empty.";

//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Title property value is empty.";
//                        }
//                    }

//                }
//                else if (checkString == "Subject should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Subject") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (checkType == 1)
//                    {
//                        if (result != "")
//                        {
//                            PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                            SortedDictionary<String, String> inf = new SortedDictionary<String, String>();
//                            inf.Add("Subject", "");
//                            stamper.MoreInfo = inf;
//                            stamper.Close();
//                            rObj.QC_Result = "Fixed";
//                            rObj.Comments = "Subject property value is not empty.";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Subject property value is empty.";
//                        }
//                    }
//                    else
//                    {
//                        if (result != "")
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Subject property value is not empty.";

//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Subject property value is empty.";
//                        }
//                    }

//                }
//                else if (checkString == "Author should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Author") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (checkType == 1)
//                    {
//                        if (result != "")
//                        {
//                            PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                            SortedDictionary<String, String> inf = new SortedDictionary<String, String>();
//                            inf.Add("Author", "");
//                            stamper.MoreInfo = inf;
//                            stamper.Close();
//                            rObj.QC_Result = "Fixed";
//                            rObj.Comments = "Author property value is not empty.";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Author property value is empty.";
//                        }
//                    }
//                    else
//                    {
//                        if (result != "")
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Author property value is not empty.";

//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Author property value is empty.";
//                        }
//                    }


//                }
//                else if (checkString == "Keywords should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Keywords") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (checkType == 1)
//                    {
//                        if (result != "")
//                        {
//                            PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                            SortedDictionary<String, String> inf = new SortedDictionary<String, String>();
//                            inf.Add("Keywords", "");
//                            stamper.MoreInfo = inf;
//                            stamper.Close();
//                            rObj.QC_Result = "Fixed";
//                            rObj.Comments = "Keywords property value is not empty.";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Keywords property value is empty.";
//                        }
//                    }
//                    else
//                    {
//                        if (result != "")
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Keywords property value is not empty.";

//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "Keywords property value is empty.";
//                        }
//                    }

//                }
//                res = qObj.SaveValidateResults(rObj);
//                reader.Close();
//                // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                string destFile = destPath;
//                System.IO.File.Copy(destFile, sourcePath, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string PDFFile_NoOfPages(RegOpsQC rObj, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {

//                sourcePath = destPath1 + rObj.Job_ID + "/Source/" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                PdfReader reader = new PdfReader(sourcePath);
//                Int64 numpages = reader.NumberOfPages;
//                rObj.QC_Result = "Pass";
//                rObj.Comments = numpages.ToString();
//                res = qObj.SaveValidateResults(rObj);
//                reader.Close();
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string PDF_Bookmarks(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                //sourcePath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                // string logfile = rObj.Job_ID + "/Destination/log.txt";
//                String fileName = sourcePath;
//                //StreamWriter outputfile = new StreamWriter(logfile);
//                if (File.Exists(fileName))
//                {
//                    PdfReader R = new PdfReader(fileName);
//                    int nothingcnt = 0;
//                    int totallinks = 0;
//                    int totallinksbroken = 0;
//                    int totallinkswork = 0;
//                    string linkTextBuilder = "";
//                    string linkReferenceBuilder = "";
//                    StringBuilder sb = new StringBuilder();
//                    StringBuilder sb1 = new StringBuilder();
//                    string workinglinks = "";
//                    string brokenlinks = "";
//                    for (int page = 1; page <= R.NumberOfPages; page++)
//                    {
//                        PdfDictionary PageDictionary = R.GetPageN(page);
//                        PdfArray Annots = PageDictionary.GetAsArray(PdfName.ANNOTS);
//                        if ((Annots == null) || (Annots.Length == 0))
//                        {
//                            nothingcnt++;
//                        }
//                        if (Annots != null)
//                        {
//                            foreach (PdfObject A in Annots.ArrayList)
//                            {
//                                //Convert the itext-specific object as a generic PDF object
//                                PdfDictionary AnnotationDictionary =
//                                (PdfDictionary)PdfReader.GetPdfObject(A);
//                                //Make sure this annotation has a link
//                                if (!AnnotationDictionary.Get(PdfName.SUBTYPE).Equals(PdfName.LINK))
//                                    continue;
//                                //Make sure this annotation has an ACTION
//                                if (AnnotationDictionary.Get(PdfName.A) == null)
//                                {
//                                    //Console.WriteLine ("no action at page no. " + page);
//                                    continue;
//                                }
//                                //Get the ACTION for the current annotation
//                                PdfDictionary AnnotationAction = AnnotationDictionary.GetAsDict(PdfName.A);
//                                // Test if it is a URI action (There are tons of other types of actions,
//                                // some of which might mimic URI, such as JavaScript,
//                                // but those need to be handled seperately)
//                                var str = AnnotationAction.Get(PdfName.S);
//                                var str1 = PdfName.GOTO;
//                                if (AnnotationAction.Get(PdfName.S).Equals(PdfName.URI))
//                                {
//                                    PdfString Link = AnnotationAction.GetAsString(PdfName.GOTO);
//                                    if (Link != null)
//                                        linkReferenceBuilder = Link.ToString();
//                                    //Get action link text : linkTextBuilder
//                                    var LinkLocation = AnnotationDictionary.GetAsArray(PdfName.RECT);
//                                    List<string> linestringlist = new List<string>();
//                                    iTextSharp.text.Rectangle rect = new iTextSharp.text.Rectangle(((PdfNumber)LinkLocation[0]).FloatValue, ((PdfNumber)LinkLocation[1]).FloatValue, ((PdfNumber)LinkLocation[2]).FloatValue, ((PdfNumber)LinkLocation[3]).FloatValue);
//                                    RenderFilter[] renderFilter = new RenderFilter[1];
//                                    renderFilter[0] = new RegionTextRenderFilter(rect);
//                                    ITextExtractionStrategy textExtractionStrategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), renderFilter);
//                                    linkTextBuilder = PdfTextExtractor.GetTextFromPage(R, page, textExtractionStrategy).Trim().Replace("..", "");
//                                    sb.AppendLine(linkTextBuilder);
//                                    workinglinks = sb.ToString();
//                                    //workinglinks = workinglinks+linkTextBuilder;
//                                    totallinks++;
//                                    totallinkswork++;
//                                }
//                                else
//                                {
//                                    PdfString Link = AnnotationAction.GetAsString(PdfName.LAUNCH);
//                                    if (Link != null)
//                                        linkReferenceBuilder = Link.ToString();
//                                    //Get action link text : linkTextBuilder
//                                    var LinkLocation = AnnotationDictionary.GetAsArray(PdfName.RECT);
//                                    List<string> linestringlist = new List<string>();
//                                    iTextSharp.text.Rectangle rect = new iTextSharp.text.Rectangle(((PdfNumber)LinkLocation[0]).FloatValue, ((PdfNumber)LinkLocation[1]).FloatValue, ((PdfNumber)LinkLocation[2]).FloatValue, ((PdfNumber)LinkLocation[3]).FloatValue);
//                                    RenderFilter[] renderFilter = new RenderFilter[1];
//                                    renderFilter[0] = new RegionTextRenderFilter(rect);
//                                    ITextExtractionStrategy textExtractionStrategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), renderFilter);
//                                    linkTextBuilder = PdfTextExtractor.GetTextFromPage(R, page, textExtractionStrategy).Trim().Replace("..", "");
//                                    sb1.AppendLine(linkTextBuilder);
//                                    brokenlinks = sb1.ToString();
//                                    totallinks++;
//                                    totallinksbroken++;
//                                }
//                            }

//                        }
//                    }
//                    if (totallinks.ToString() != "0")
//                    {
//                        string totalLinks = "Total number of links Count - " + totallinks.ToString();
//                        string workingLinksCount = "Total number of Working Links Count - " + totallinkswork.ToString();
//                        string workingLinks = "List of working links Names - " + workinglinks;
//                        string totalbrokenlinks = string.Empty;
//                        string brokenLinksList = string.Empty;
//                        if (totallinksbroken.ToString() != "0")
//                        {
//                            totalbrokenlinks = "Total number of Broken links Count - " + totallinksbroken.ToString();
//                            brokenLinksList = "List of broken links Names - " + brokenlinks;
//                            rObj.Comments = totalLinks + "," + workingLinksCount + "," + totalbrokenlinks + "," + brokenLinksList;
//                        }
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = totalLinks + "," + workingLinksCount;
//                        res = qObj.SaveValidateResults(rObj);
//                        R.Close();
//                        string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                        System.IO.File.Copy(destPath, destFile, true);
//                    }
//                    else
//                    {
//                        string result = "No bookmarks found in this file.";
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = result;
//                        res = qObj.SaveValidateResults(rObj);
//                        R.Close();
//                        string destFile1 = System.IO.Path.Combine(path, rObj.File_Name);
//                        System.IO.File.Copy(destPath, destFile1, true);
//                    }

//                }
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string PDFMagnification(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                // sourcePath = destPath1 + rObj.Job_ID + "/Source/" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                Document document = new Document();
//                FileInfo f1 = new FileInfo(destPath);
//                string str = f1.Name;
//                string str1 = f1.Extension;
//                string tempPath = destPath;

//                str = f1.Name;
//                str1 = f1.Extension;
//                //delete the existing file with same name
//                if (File.Exists(tempPath))
//                {
//                    File.Delete(tempPath);
//                }
//                string src = sourcePath;
//                string dest = destPath;

//                PdfReader pdf = new PdfReader(src);

//                PdfStamper stamper = new PdfStamper(pdf, new FileStream(dest, FileMode.Create));

//                //setting zoom level to 100%
//                PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, pdf.GetPageSize(1).Height, 1f);
//                PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, stamper.Writer);
//                stamper.Writer.SetOpenAction(action);

//                stamper.Close();
//                //open the newly created file
//                // System.Diagnostics.Process.Start(dest);
//                //rObj.QC_Result = "Failed";
//                rObj.QC_Result = "Fixed";
//                rObj.Comments = "Magnification has been set to default.";
//                res = qObj.SaveValidateResults(rObj);
//                pdf.Close();
//               // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                System.IO.File.Copy(dest, src, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string PDFPageLayout(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                //sourcePath = destPath1 + rObj.Job_ID + "/Source/" + rObj.File_Name;
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                Document document = new Document();
//                FileInfo f1 = new FileInfo(destPath);
//                string str = f1.Name;
//                string str1 = f1.Extension;
//                string tempPath = destPath;

//                str = f1.Name;
//                str1 = f1.Extension;

//                //delete the existing file with same name
//                if (File.Exists(tempPath))
//                {
//                    File.Delete(tempPath);
//                }
//                string src = sourcePath;
//                string dest = destPath;

//                PdfReader pdf = new PdfReader(src);

//                //setting page layout to single page
//                pdf.AddViewerPreference(PdfName.PRINTSCALING, PdfName.DEFAULT);
//                pdf.AddViewerPreference(PdfName.PAGELAYOUT, PdfName.SINGLEPAGE);
//                PdfStamper stamper = new PdfStamper(pdf, new FileStream(dest, FileMode.Create));

//                stamper.Close();
//                pdf.Close();
//                //open the newly created file
//                // System.Diagnostics.Process.Start(dest);
//                //rObj.QC_Result = "Failed";
//                rObj.QC_Result = "Fixed";
//                rObj.Comments = "Page layout has been set to default.";
//                res = qObj.SaveValidateResults(rObj);
//               // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                System.IO.File.Copy(dest, src, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string DeleteExternalHyperlinks(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                String fileName = sourcePath;
//                string pageNumbers = "";
//                if (File.Exists(fileName))
//                {
//                    int num1 = 0;
//                    int num2 = 0;
//                    int num3 = 0;
//                    PdfReader reader = new PdfReader(fileName);
//                    List<string> stringList1 = new List<string>();
//                    List<string> stringList2 = new List<string>();
//                    StringBuilder stringBuilder1 = new StringBuilder();
//                    StringBuilder stringBuilder2 = new StringBuilder();
//                    for (num1 = 1; num1 <= reader.NumberOfPages; ++num1)
//                    {
//                        PdfArray asArray1 = reader.GetPageN(num1).GetAsArray(PdfName.ANNOTS);
//                        if (asArray1 == null || asArray1.Length == 0)
//                            ++num2;
//                        if (asArray1 != null)
//                        {
//                            foreach (PdfObject array in asArray1.ArrayList)
//                            {
//                                PdfDictionary pdfObject = (PdfDictionary)PdfReader.GetPdfObject(array);
//                                if (pdfObject.Get(PdfName.SUBTYPE).Equals((object)PdfName.LINK) && pdfObject.Get(PdfName.A) != null)
//                                {
//                                    PdfDictionary asDict = pdfObject.GetAsDict(PdfName.A);
//                                    string str5;
//                                    if (asDict.Get(PdfName.S).Equals((object)PdfName.URI))
//                                    {
//                                        PdfArray asArray2 = pdfObject.GetAsArray(PdfName.RECT);
//                                        ITextExtractionStrategy strategy = (ITextExtractionStrategy)new FilteredTextRenderListener((ITextExtractionStrategy)new LocationTextExtractionStrategy(), new RenderFilter[1]
//                                        {
//                    (RenderFilter) new RegionTextRenderFilter(new iTextSharp.text.Rectangle(((PdfNumber) asArray2[0]).FloatValue, ((PdfNumber) asArray2[1]).FloatValue, ((PdfNumber) asArray2[2]).FloatValue, ((PdfNumber) asArray2[3]).FloatValue))
//                                        });
//                                        str5 = PdfTextExtractor.GetTextFromPage(reader, num1, strategy).Trim().Replace("..", "");
//                                        //Removing Link
//                                        asDict.Remove(PdfName.URI);
//                                        //Adding Page numbers
//                                        if (pageNumbers == "")
//                                        {
//                                            pageNumbers = num1.ToString() + ",";
//                                        }
//                                        else if ((!pageNumbers.Contains(num1.ToString() + ",")))
//                                            pageNumbers = pageNumbers + num1.ToString() + ",";

//                                        ++num3;
//                                    }
//                                }

//                            }
//                        }
//                    }
//                    PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                    stamper.Close();
//                    if (num3 > 0)
//                    {
//                        //rObj.QC_Result = "Failed";
//                        rObj.QC_Result = "Fixed";
//                        if (pageNumbers != "")
//                        {
//                            pageNumbers = pageNumbers.Trim(',');
//                            rObj.Comments = pageNumbers;
//                        }
//                        //rObj.Comments = "Deleted external Hyperlinks.";
//                        res = qObj.SaveValidateResults(rObj);
//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "There is no external Hyperlinks.";

//                        res = qObj.SaveValidateResults(rObj);
//                    }
//                    reader.Close();
//                    System.IO.File.Copy(destPath, sourcePath, true);
//                }
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string ListBrokenInternalHyperLinks(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            string pageNumbers = "";
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                String fileName = sourcePath;
//                if (File.Exists(fileName))
//                {
//                    int num1 = 0;
//                    int num2 = 0;
//                    int num3 = 0;
//                    int num4 = 0;
//                    int num5 = 0;
//                    List<string> stringList1 = new List<string>();
//                    List<string> stringList2 = new List<string>();
//                    string str2 = "";
//                    StringBuilder stringBuilder1 = new StringBuilder();
//                    StringBuilder stringBuilder2 = new StringBuilder();
//                    string str3 = "";
//                    string str4 = "";
//                    PdfReader reader = new PdfReader(fileName);
//                    for (num1 = 1; num1 <= reader.NumberOfPages; ++num1)
//                    {
//                        PdfArray asArray1 = reader.GetPageN(num1).GetAsArray(PdfName.ANNOTS);
//                        if (asArray1 == null || asArray1.Length == 0)
//                            ++num2;
//                        if (asArray1 != null)
//                        {
//                            foreach (PdfObject array in asArray1.ArrayList)
//                            {
//                                PdfDictionary pdfObject = (PdfDictionary)PdfReader.GetPdfObject(array);
//                                if (pdfObject.Get(PdfName.SUBTYPE).Equals((object)PdfName.LINK) && pdfObject.Get(PdfName.A) != null)
//                                {
//                                    PdfDictionary asDict = pdfObject.GetAsDict(PdfName.A);
//                                    List<string> stringList3;
//                                    if (asDict.Get(PdfName.S).Equals((object)PdfName.GOTO))
//                                    {
//                                        PdfString asString = asDict.GetAsString(PdfName.GOTO);
//                                        if (asString != null)
//                                            str2 = asString.ToString();
//                                        PdfArray asArray2 = pdfObject.GetAsArray(PdfName.RECT);
//                                        stringList3 = new List<string>();
//                                        ITextExtractionStrategy strategy = (ITextExtractionStrategy)new FilteredTextRenderListener((ITextExtractionStrategy)new LocationTextExtractionStrategy(), new RenderFilter[1]
//                                        {
//                    (RenderFilter) new RegionTextRenderFilter(new iTextSharp.text.Rectangle(((PdfNumber) asArray2[0]).FloatValue, ((PdfNumber) asArray2[1]).FloatValue, ((PdfNumber) asArray2[2]).FloatValue, ((PdfNumber) asArray2[3]).FloatValue))
//                                        });
//                                        string str5 = PdfTextExtractor.GetTextFromPage(reader, num1, strategy).Trim().Replace("..", "");
//                                        stringBuilder1.AppendLine(str5 + " - <font color='green'> Passed </font></br>");
//                                        str3 = stringBuilder1.ToString();
//                                        ++num3;
//                                        ++num5;
//                                    }
//                                    else
//                                    {
//                                        PdfString asString = asDict.GetAsString(PdfName.LAUNCH);
//                                        if (asString != null)
//                                            str2 = asString.ToString();
//                                        PdfArray asArray2 = pdfObject.GetAsArray(PdfName.RECT);
//                                        stringList3 = new List<string>();
//                                        ITextExtractionStrategy strategy = (ITextExtractionStrategy)new FilteredTextRenderListener((ITextExtractionStrategy)new LocationTextExtractionStrategy(), new RenderFilter[1]
//                                        {
//                    (RenderFilter) new RegionTextRenderFilter(new iTextSharp.text.Rectangle(((PdfNumber) asArray2[0]).FloatValue, ((PdfNumber) asArray2[1]).FloatValue, ((PdfNumber) asArray2[2]).FloatValue, ((PdfNumber) asArray2[3]).FloatValue))
//                                        });
//                                        string str5 = PdfTextExtractor.GetTextFromPage(reader, num1, strategy).Trim().Replace("..", "");
//                                        stringBuilder2.AppendLine(str5 + "</br>");
//                                        str4 = stringBuilder2.ToString();
//                                        pageNumbers = pageNumbers + "," + stringBuilder2.ToString() + "- PageNo: " + num1.ToString();

//                                        ++num3;
//                                        ++num4;
//                                    }
//                                }
//                            }
//                        }
//                    }
//                    //StreamWriter streamWriter = new StreamWriter("d:/log.html");
//                    //streamWriter.WriteLine("<b>List of broken internal hyperlinks</b></br>");
//                    //streamWriter.WriteLine("</br>");
//                    //streamWriter.WriteLine(str4);
//                    PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                    stamper.Close();
//                    //if (str4!="")
//                    if (pageNumbers != "")
//                    {
//                        rObj.QC_Result = "Failed";
//                        //rObj.Comments = str4;
//                        pageNumbers = pageNumbers.Trim(',');
//                        rObj.Comments = pageNumbers;
//                        res = qObj.SaveValidateResults(rObj);
//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "There is no broken internal hyperlinks.";
//                        res = qObj.SaveValidateResults(rObj);
//                    }
//                    reader.Close();
//                   // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                    System.IO.File.Copy(destPath, sourcePath, true);
//                }
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string DeleteAllDeadLinks(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            string pageNumbers = "";
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                String fileName = sourcePath;
//                if (File.Exists(fileName))
//                {
//                    int num1 = 0;
//                    int num2 = 0;
//                    int num3 = 0;
//                    int num4 = 0;
//                    int num5 = 0;
//                    List<string> stringList1 = new List<string>();
//                    List<string> stringList2 = new List<string>();
//                    string str2 = "";
//                    StringBuilder stringBuilder1 = new StringBuilder();
//                    StringBuilder stringBuilder2 = new StringBuilder();
//                    string str3 = "";
//                    string str4 = "";
//                    PdfReader reader = new PdfReader(fileName);
//                    for (num1 = 1; num1 <= reader.NumberOfPages; ++num1)
//                    {
//                        PdfArray asArray1 = reader.GetPageN(num1).GetAsArray(PdfName.ANNOTS);
//                        if (asArray1 == null || asArray1.Length == 0)
//                            ++num2;
//                        if (asArray1 != null)
//                        {
//                            foreach (PdfObject array in asArray1.ArrayList)
//                            {
//                                PdfDictionary pdfObject = (PdfDictionary)PdfReader.GetPdfObject(array);
//                                if (pdfObject.Get(PdfName.SUBTYPE).Equals((object)PdfName.LINK) && pdfObject.Get(PdfName.A) != null)
//                                {
//                                    string str5 = string.Empty;
//                                    PdfDictionary asDict = pdfObject.GetAsDict(PdfName.A);
//                                    if (asDict.Get(PdfName.S).Equals((object)PdfName.GOTO))
//                                    {
//                                        PdfString asString = asDict.GetAsString(PdfName.GOTO);
//                                        if (asString != null)
//                                            str2 = asString.ToString();
//                                        PdfArray asArray2 = pdfObject.GetAsArray(PdfName.RECT);
//                                        ITextExtractionStrategy strategy = (ITextExtractionStrategy)new FilteredTextRenderListener((ITextExtractionStrategy)new LocationTextExtractionStrategy(), new RenderFilter[1]
//                                        {
//                    (RenderFilter) new RegionTextRenderFilter(new iTextSharp.text.Rectangle(((PdfNumber) asArray2[0]).FloatValue, ((PdfNumber) asArray2[1]).FloatValue, ((PdfNumber) asArray2[2]).FloatValue, ((PdfNumber) asArray2[3]).FloatValue))
//                                        });
//                                        str5 = PdfTextExtractor.GetTextFromPage(reader, num1, strategy).Trim().Replace("..", "");
//                                    }
//                                    else
//                                    {
//                                        PdfString asString = asDict.GetAsString(PdfName.LAUNCH);
//                                        if (asString != null)
//                                            str2 = asString.ToString();
//                                        PdfArray asArray2 = pdfObject.GetAsArray(PdfName.RECT);
//                                        ITextExtractionStrategy strategy = (ITextExtractionStrategy)new FilteredTextRenderListener((ITextExtractionStrategy)new LocationTextExtractionStrategy(), new RenderFilter[1]
//                                        {
//                    (RenderFilter) new RegionTextRenderFilter(new iTextSharp.text.Rectangle(((PdfNumber) asArray2[0]).FloatValue, ((PdfNumber) asArray2[1]).FloatValue, ((PdfNumber) asArray2[2]).FloatValue, ((PdfNumber) asArray2[3]).FloatValue))
//                                        });
//                                        str5 = PdfTextExtractor.GetTextFromPage(reader, num1, strategy).Trim().Replace("..", "");
//                                    }
//                                    reader.GetPageN(num1);
//                                    foreach (Match match in new Regex("(?<url>(http:|https:[/][/])([a-z]|[A-Z]|[0-9]|[/.]|[~])*)", RegexOptions.IgnoreCase).Matches(PdfTextExtractor.GetTextFromPage(reader, num1)))
//                                    {

//                                        HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(match.Value);
//                                        httpWebRequest.Method = "HEAD";
//                                        try
//                                        {
//                                            httpWebRequest.GetResponse().Close();
//                                            stringList1.Add(match.Value);
//                                        }
//                                        catch
//                                        {
//                                            stringList2.Add(match.Value);
//                                            //Removing Link
//                                            asDict.Remove(PdfName.URI);
//                                            if (pageNumbers == "")
//                                            {
//                                                pageNumbers = num1.ToString() + ",";
//                                            }
//                                            else if ((!pageNumbers.Contains(num1.ToString() + ",")))
//                                                pageNumbers = pageNumbers + num1.ToString() + ",";
//                                            str5.Remove(0);
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                    }
//                    if (stringList2.Count > 0)
//                    {
//                        PdfStamper stamper = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                        stamper.Close();
//                        //rObj.QC_Result = "Failed";
//                        rObj.QC_Result = "Fixed";
//                        if (pageNumbers != "")
//                        {
//                            pageNumbers = pageNumbers.Trim(',');
//                            rObj.Comments = pageNumbers;
//                        }
//                        //rObj.Comments = "Deleted all dead links.";
//                        res = qObj.SaveValidateResults(rObj);
//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "There is no dead links.";
//                        res = qObj.SaveValidateResults(rObj);
//                    }
//                    reader.Close();
//                   // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                    System.IO.File.Copy(destPath, sourcePath, true);
//                }
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string BookMarksZoomLevelSetToInheritZoom(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                String fileName = sourcePath;
//                if (File.Exists(fileName))
//                {
//                    PdfReader reader = new PdfReader(sourcePath);
//                    PdfStamper pdfStamper1 = new PdfStamper(reader, new FileStream(destPath, FileMode.Create));
//                    int n = reader.NumberOfPages;
//                    for (int i = 1; i <= n; i++)
//                    {

//                        PdfDictionary pageDic = reader.GetPageN(i);
//                        PdfArray annots = pageDic.GetAsArray(PdfName.ANNOTS);
//                        if (annots != null)
//                        {
//                            for (int j = 0; j < annots.Size; j++)
//                            {
//                                PdfDictionary annotation = annots.GetAsDict(j);
//                                PdfDictionary annotationAction = annotation.GetAsDict(PdfName.A);
//                                if (annotationAction == null)
//                                    continue;
//                                PdfName actionType = annotationAction.GetAsName(PdfName.S);

//                                PdfArray d = null;
//                                if (PdfName.GOTO.Equals(actionType))
//                                    d = annotationAction.GetAsArray(PdfName.D);
//                                else if (PdfName.LINK.Equals(actionType))
//                                    d = annotation.GetAsArray(PdfName.DEST);
//                                if (d == null)
//                                    continue;
//                                if (d.Size == 5 && PdfName.XYZ.Equals(d.GetAsName(1)))
//                                {
//                                    d[4] = new PdfNumber(0);
//                                }
//                            }
//                        }
//                    }
//                    //rObj.QC_Result = "Failed";
//                    rObj.QC_Result = "Fixed";
//                    rObj.Comments = "Zoom level set to inherit Zoom.";
//                    res = qObj.SaveValidateResults(rObj);
//                    pdfStamper1.Close();
//                    reader.Close();
//                   // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                    System.IO.File.Copy(destPath, sourcePath, true);
//                }
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string PDFNavigationTabSetToPageOnly(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                // sourcePath = destPath1 + rObj.Job_ID + "/Source/" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                Document document = new Document();
//                FileInfo f1 = new FileInfo(destPath);
//                string str = f1.Name;
//                string str1 = f1.Extension;
//                string tempPath = destPath;

//                str = f1.Name;
//                str1 = f1.Extension;
//                //delete the existing file with same name
//                if (File.Exists(tempPath))
//                {
//                    File.Delete(tempPath);
//                }
//                string src = sourcePath;
//                string dest = destPath;

//                PdfReader pdf = new PdfReader(src);
//                int pageCount = pdf.NumberOfPages;
//                PdfStamper stamper = new PdfStamper(pdf, new FileStream(dest, FileMode.Create));
//                //if (pageCount < 4)
//                if (pageCount > 0)
//                {

//                    pdf.AddViewerPreference(PdfName.PRINTSCALING, PdfName.DEFAULT);
//                    pdf.AddViewerPreference(PdfName.PAGELAYOUT, PdfName.SINGLEPAGE);
//                    /* the aboe two lines or the below line is working for page navigation tab to page only*/
//                    //pdf.AddViewerPreference(PdfName.NAVIGATIONPANE, PdfName.PAGE);
//                    //setting zoom level to 100%
//                    PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, pdf.GetPageSize(1).Height, 1f);
//                    PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, stamper.Writer);
//                    stamper.Writer.SetOpenAction(action);

//                    rObj.QC_Result = "Fixed";
//                    rObj.Comments = "Navigation tab has been set to Page only.";
//                }
//                else
//                {
//                    rObj.QC_Result = "Passed";
//                    rObj.Comments = "The file has more than 4 pages.";
//                }
//                stamper.Close();
//                res = qObj.SaveValidateResults(rObj);
//                pdf.Close();
//               // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                System.IO.File.Copy(destPath, sourcePath, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string CheckNoOfHyperLinks(RegOpsQC rObj, string bookmarkfile1, string destPath)
//        {
//            //string str1 = args[0];
//            string res = string.Empty;
//            rObj.CHECK_START_TIME = DateTime.Now;
//            String fileName = bookmarkfile1 + "\\" + rObj.File_Name;
//            //StreamWriter streamWriter = new StreamWriter("d:/log.html");
//            int num1 = 0;
//            try
//            {
//                if (!System.IO.File.Exists(fileName))
//                    return "File not found";
//                PdfReader reader = new PdfReader(fileName);
//                int num2 = 0;
//                int num3 = 0;
//                int num4 = 0;
//                int num5 = 0;
//                List<string> stringList1 = new List<string>();
//                List<string> stringList2 = new List<string>();
//                string str2 = "";
//                StringBuilder stringBuilder1 = new StringBuilder();
//                StringBuilder stringBuilder2 = new StringBuilder();
//                string str3 = "";
//                string str4 = "";
//                for (num1 = 1; num1 <= reader.NumberOfPages; ++num1)
//                {
//                    PdfArray asArray1 = reader.GetPageN(num1).GetAsArray(PdfName.ANNOTS);
//                    if (asArray1 == null || asArray1.Length == 0)
//                        ++num2;
//                    if (asArray1 != null)
//                    {
//                        foreach (PdfObject array in asArray1.ArrayList)
//                        {
//                            PdfDictionary pdfObject = (PdfDictionary)PdfReader.GetPdfObject(array);
//                            if (pdfObject.Get(PdfName.SUBTYPE).Equals((object)PdfName.LINK) && pdfObject.Get(PdfName.A) != null)
//                            {
//                                PdfDictionary asDict = pdfObject.GetAsDict(PdfName.A);
//                                List<string> stringList3;
//                                if (asDict.Get(PdfName.S).Equals((object)PdfName.GOTO))
//                                {
//                                    PdfString asString = asDict.GetAsString(PdfName.GOTO);
//                                    if (asString != null)
//                                        str2 = asString.ToString();
//                                    PdfArray asArray2 = pdfObject.GetAsArray(PdfName.RECT);
//                                    stringList3 = new List<string>();
//                                    ITextExtractionStrategy strategy = (ITextExtractionStrategy)new FilteredTextRenderListener((ITextExtractionStrategy)new LocationTextExtractionStrategy(), new RenderFilter[1]
//                                    {
//                    (RenderFilter) new RegionTextRenderFilter(new iTextSharp.text.Rectangle(((PdfNumber) asArray2[0]).FloatValue, ((PdfNumber) asArray2[1]).FloatValue, ((PdfNumber) asArray2[2]).FloatValue, ((PdfNumber) asArray2[3]).FloatValue))
//                                    });
//                                    string str5 = PdfTextExtractor.GetTextFromPage(reader, num1, strategy).Trim().Replace("..", "");
//                                    str3 = stringBuilder1.ToString();
//                                    ++num3;
//                                    ++num5;
//                                }
//                                else
//                                {
//                                    PdfString asString = asDict.GetAsString(PdfName.LAUNCH);
//                                    if (asString != null)
//                                        str2 = asString.ToString();
//                                    PdfArray asArray2 = pdfObject.GetAsArray(PdfName.RECT);
//                                    stringList3 = new List<string>();
//                                    ITextExtractionStrategy strategy = (ITextExtractionStrategy)new FilteredTextRenderListener((ITextExtractionStrategy)new LocationTextExtractionStrategy(), new RenderFilter[1]
//                                    {
//                    (RenderFilter) new RegionTextRenderFilter(new iTextSharp.text.Rectangle(((PdfNumber) asArray2[0]).FloatValue, ((PdfNumber) asArray2[1]).FloatValue, ((PdfNumber) asArray2[2]).FloatValue, ((PdfNumber) asArray2[3]).FloatValue))
//                                    });
//                                    string str5 = PdfTextExtractor.GetTextFromPage(reader, num1, strategy).Trim().Replace("..", "");
//                                    str4 = stringBuilder2.ToString();
//                                    ++num3;
//                                    ++num4;
//                                }
//                            }
//                        }
//                    }
//                }
//                for (num1 = 1; num1 <= reader.NumberOfPages; ++num1)
//                {
//                    reader.GetPageN(num1);
//                    foreach (Match match in new Regex("(?<url>(http:|https:[/][/])([a-z]|[A-Z]|[0-9]|[/.]|[~])*)", RegexOptions.IgnoreCase).Matches(PdfTextExtractor.GetTextFromPage(reader, num1)))
//                    {
//                        try
//                        {
//                            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(match.Value);
//                            httpWebRequest.Method = "HEAD";
//                            try
//                            {
//                                httpWebRequest.GetResponse().Close();
//                                stringList1.Add(match.Value);
//                            }
//                            catch
//                            {
//                                stringList2.Add(match.Value);
//                            }
//                        }
//                        catch
//                        {
//                            //streamWriter.WriteLine();
//                            //streamWriter.WriteLine("Error while parsing page no: " + num1.ToString());
//                        }
//                    }
//                }
//                int num6 = stringList1.Count + stringList2.Count;
//                int totalNoOfLinks = 0;
//                totalNoOfLinks = (num3 + num6);
//                //return (num3 + num6).ToString();

//                //rObj.QC_Result = "";

//                if (totalNoOfLinks > 0)
//                {
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = "Total hyperlinks: " + totalNoOfLinks.ToString();
//                }
//                else
//                {
//                    rObj.QC_Result = "Passed";
//                    rObj.Comments = "No hyperlinks existed in the file.";
//                }
//                res = qObj.SaveValidateResults(rObj);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }

//        }

//        public string PDFPageLayoutAndMagnification(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                //sourcePath = destPath1 + rObj.Job_ID + "/Source/" + rObj.File_Name;
//                sourcePath = path + "//" + rObj.File_Name;
//                // destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                Document document = new Document();
//                FileInfo f1 = new FileInfo(destPath);
//                string str = f1.Name;
//                string str1 = f1.Extension;
//                string tempPath = destPath;

//                str = f1.Name;
//                str1 = f1.Extension;

//                //delete the existing file with same name
//                if (File.Exists(tempPath))
//                {
//                    File.Delete(tempPath);
//                }
//                string src = sourcePath;
//                string dest = destPath;

//                PdfReader pdf = new PdfReader(src);

//                //setting page layout to single page
//                pdf.AddViewerPreference(PdfName.PRINTSCALING, PdfName.DEFAULT);
//                pdf.AddViewerPreference(PdfName.PAGELAYOUT, PdfName.SINGLEPAGE);

//                PdfStamper stamper = new PdfStamper(pdf, new FileStream(dest, FileMode.Create));

//                //setting zoom level to 100%
//                PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, pdf.GetPageSize(1).Height, 1f);
//                PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, stamper.Writer);

//                stamper.Writer.SetOpenAction(action);

//                stamper.Close();
//                pdf.Close();
//                //open the newly created file
//                // System.Diagnostics.Process.Start(dest);
//                //rObj.QC_Result = "Failed";
//                rObj.QC_Result = "Fixed";
//                rObj.Comments = "Page layout and magnification set to default.";
//                res = qObj.SaveValidateResults(rObj);
//               // string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                System.IO.File.Copy(dest, sourcePath, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        public string PDFFile_VerifyProperties(RegOpsQC rObj, string path, string checkString, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                // sourcePath = destPath1 + rObj.Job_ID + "/Source/" + rObj.File_Name;
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                PdfReader reader = new PdfReader(sourcePath);
//                IDictionary<String, String> metadic = reader.Info;
//                var dic = from m in metadic
//                          select m;
//                string result = string.Empty;

//                if (checkString == "Document properties should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        //Console.WriteLine(d.Key + ": " + d.Value);
//                        //result = result + " , " + d.Key + ": " + d.Value;
//                        if ((d.Key == "Title" || d.Key == "Subject" || d.Key == "Author" || d.Key == "Keywords") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (result != "")
//                    {
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Default properties are not empty.";

//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "Default properties are empty.";
//                    }

//                }
//                else if (checkString == "Title should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Title") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (result != "")
//                    {
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Title property value is not empty.";

//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "Title property value is empty.";
//                    }
//                }
//                else if (checkString == "Subject should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Subject") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (result != "")
//                    {
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Subject property value is not empty";

//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "Subject property value is empty.";
//                    }
//                }
//                else if (checkString == "Author should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Author") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (result != "")
//                    {
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Author property value is not empty.";

//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "Author property value is empty";
//                    }

//                }
//                else if (checkString == "Keywords should be blank")
//                {
//                    foreach (var d in dic)
//                    {
//                        if ((d.Key == "Keywords") && d.Value != "")
//                            result = result + " , " + d.Key + ": " + d.Value;
//                    }
//                    if (result != "")
//                    {
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Keywords property value is not empty.";

//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "Keywords property value is empty.";
//                    }
//                }

//                res = qObj.SaveValidateResults(rObj);
//                reader.Close();
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }
//        public string EnableOCR1(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            // step 0: set minimum page size
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                using (Bytescout.PDFExtractor.TextExtractor extractor = new Bytescout.PDFExtractor.TextExtractor())
//                {
//                    extractor.RegistrationName = "demo";
//                    extractor.RegistrationKey = "demo";

//                    // setup OCR
//                    extractor.OCRMode = OCRMode.Auto;
//                    // extractor.OCRLanguageDataFolder = ocrLanguageDataFolder;
//                    extractor.OCRLanguage = "eng";
//                    extractor.OCRResolution = 300;
//                    extractor.LoadDocumentFromFile(destPath);
//                }
//                rObj.QC_Result = "Fixed";
//                rObj.Comments = "OCR is enabled.";
//                res = qObj.SaveValidateResults(rObj);
//                System.IO.File.Copy(destPath, sourcePath, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        //25. Check for Blank pages
//        public string checkForBlankPdfPages(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            // step 0: set minimum page size
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                // step 1: create new reader
//                var r = new PdfReader(sourcePath);
//                //var raf = new RandomAccessFileOrArray(sourcePath);
//                var document = new Document(r.GetPageSizeWithRotation(1));
//                int count = 0;
//              // step 3: we open the document
//                document.Open();
//                // step 4: we add content
//                PdfImportedPage page = null;
//                for (int pages = 1; pages <= r.NumberOfPages; pages++)
//                {
//                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
//                    string currentText = PdfTextExtractor.GetTextFromPage(r, pages, strategy);

//                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
//                    if (currentText.Length == 0)
//                    {
//                        count = count + 1;
//                    }
//                }
//                if (count == 0)
//                {
//                    rObj.QC_Result = "Passed";
//                    rObj.Comments = "No blank pages existed in given document";
//                }
//                else
//                {
//                    rObj.QC_Result = "Failed";
//                    Console.WriteLine("Number of blank pages={0}", count);
//                    rObj.Comments = "Blank pages exists in the given document.";
//                }
//                r.Close();
//                res = qObj.SaveValidateResults(rObj);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        //26.Remove blank pages
//        public string removeBlankPdfPages(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                // step 0: set minimum page size
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                //string src = sourcePath;
//                //string dest = System.IO.Path.GetDirectoryName(destPath);
//                string pagescount = string.Empty;
//                // step 1: create new reader
//                var r = new PdfReader(sourcePath);
//                var raf = new RandomAccessFileOrArray(sourcePath);
//                var document = new Document(r.GetPageSizeWithRotation(1));
//             //   document.Open();
//                int count = 0;
//                int numberofpages = r.NumberOfPages;
//                for (int pages = 1; pages <= numberofpages; pages++)
//                {
//                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
//                    string currentText = PdfTextExtractor.GetTextFromPage(r, pages, strategy);

//                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
//                    if (currentText.Length == 0)
//                    {
//                        count = count + 1;
//                        //pagescount = pagescount + "," + pages;
//                    }
//                    else if (currentText.Length > 0)
//                    {
//                        //count = count + 1;
//                        pagescount = pagescount + "," + pages;
//                    }
//                }

//                string src = sourcePath;
//                string dest = destPath;
                               
//                    if (rObj.Check_Type == 1)
//                    {
//                        if (count == 0)
//                        {                        
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "No blank pages are present in given document.";
//                        }
//                        else if (pagescount != "" && count != 0)
//                        {
//                            string mypages = pagescount.Substring(1);
//                            r.SelectPages(mypages);
//                            PdfStamper stamper = new PdfStamper(r, new FileStream(dest, FileMode.Create));
//                            stamper.Close();
//                            rObj.QC_Result = "Fixed";
//                            rObj.Comments = count + " blank page(s) are removed.";
//                        }                       
//                    }
//                    else
//                    {
//                        if (count == 0)
//                        {                            
//                            rObj.QC_Result = "Passed";
//                            rObj.Comments = "No blank pages are present in given document.";
//                        }
//                        if(pagescount != "" && count != 0)
//                        {                                                   
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = count + " blank page(s) are exists in the given document.";
//                        }                        
//                    }                              
//                r.Close();
//                raf.Close();
//                document.Close();
//                System.IO.File.Copy(dest, sourcePath, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res = qObj.SaveValidateResults(rObj); ;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                rObj.Job_Status = "Failed";
//                rObj.QC_Result = "Error";
//                rObj.Comments = "Technical error: " + ex.Message;
//                qObj.SaveValidateResults(rObj);
//                return "Error";
//            }
//        }

//        //35.Check for Track Changes
//        public string checkForannotations(RegOpsQC rObj, string path, string destPath)
//        {
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                string pagescount = string.Empty;
//                //Bytes will hold our final PDFs
//                byte[] bytes;
//                int count = 0;
//                using (var ms = new MemoryStream())
//                {
//                    using (var reader = new PdfReader(sourcePath))
//                    {
//                        using (PdfStamper stamper = new PdfStamper(reader, ms))
//                        {
//                            for (int i = 1; i <= reader.NumberOfPages; i++)
//                            {
//                                // get a page a PDF page
//                                PdfDictionary page = reader.GetPageN(i);
//                                // get all the annotations of page i
//                                PdfArray annotationsArray = page.GetAsArray(PdfName.ANNOTS);

//                                // if page does not have annotations
//                                if (annotationsArray == null)
//                                {
//                                    count = count + 1;
//                                    continue;
//                                }
//                                // for each annotation
//                                for (int j = 0; j < annotationsArray.Size; j++)
//                                {
//                                    // for current annotation
//                                    PdfDictionary currentAnnotation = annotationsArray.GetAsDict(j);

//                                    PdfDictionary annotationAction = currentAnnotation.GetAsDict(PdfName.AA);
//                                    if (annotationAction == null)
//                                    {
//                                        annotationsArray.Remove(j);
//                                        Console.Write("Removed annotation {0} with no action from page {1}\n", j, i);
//                                        //if(!pagescount.Contains(","+ i))
//                                        //    pagescount = pagescount + "," + i;
//                                        if (pagescount == "")
//                                            pagescount = i.ToString() + ",";
//                                        else if ((!pagescount.Contains(i.ToString() + ",")))
//                                            pagescount = pagescount + i.ToString() + ",";
//                                    }
//                                }
//                            }
//                            if (count == reader.NumberOfPages)
//                            {
//                                rObj.QC_Result = "Passed";
//                                rObj.Comments = "Track changes are not present in given file.";
//                            }
//                            else
//                            {
//                                rObj.QC_Result = "Fixed";
//                                rObj.Comments = "Track changes are removed in the following page numbers :"+pagescount.Trim(',');
//                            }                                                                                 
//                        }
//                    }
//                    bytes = ms.ToArray();
//                }
//                File.WriteAllBytes(destPath, bytes);
//                res = qObj.SaveValidateResults(rObj);
//                //reader.Close();
//              //  string destFile = System.IO.Path.Combine(path, rObj.File_Name);
//                System.IO.File.Copy(destPath, sourcePath, true);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }
//        //36.Check for consistency between subheadings and table of contents entries
//        public string tOCSubheadings(RegOpsQC rObj, string path, string destPath)

//        {
//            //Setup some variables to be used later
//            PdfReader R = default(PdfReader);
//            int PageCount = 0;
//            string res = string.Empty;
//            try
//            {
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                //Open our reader
//                R = new PdfReader(sourcePath);
//                //Get the page cont
//                PageCount = R.NumberOfPages;
//                Console.WriteLine("Page Count= " + PageCount);

//                //Loop through each page
//                //for (int i = 1; i <= PageCount; i++)
//                //{
//                //Get the current page
//                PdfDictionary PageDictionary = R.GetPageN(1);
//                //Get all of the annotations for the current page
//                PdfArray Annots = PageDictionary.GetAsArray(PdfName.ANNOTS);
//                //Make sure we have something
//                if ((Annots == null) || (Annots.Length == 0))
//                {
//                    Console.WriteLine("nothing");
//                }
//                //Loop through each annotation
//                if (Annots != null)
//                {
//                    Console.WriteLine("ANNOTS Not Null" + Annots[0]);
//                    foreach (PdfObject A in Annots.ArrayList)
//                    {
//                        //Convert the itext-specific object as a generic PDF object
//                        PdfDictionary AnnotationDictionary = (PdfDictionary)PdfReader.GetPdfObject(A);
//                        //Make sure this annotation has a link
//                        if (!AnnotationDictionary.Get(PdfName.SUBTYPE).Equals(PdfName.LINK))
//                            continue;
//                        //Make sure this annotation has an ACTION
//                        if (AnnotationDictionary.Get(PdfName.A) == null)
//                            continue;
//                        if (AnnotationDictionary.Get(PdfName.A) != null)
//                        {
//                            Console.WriteLine("ACTION Not Null");
//                        }
//                        //Get the ACTION for the current annotation
//                        PdfDictionary AnnotationAction = AnnotationDictionary.GetAsDict(PdfName.A);

//                        // Test if it is a URI action (There are tons of other types of actions,
//                        // some of which might mimic URI, such as JavaScript,
//                        // but those need to be handled seperately)
//                        if (AnnotationAction.Get(PdfName.S).Equals(PdfName.URI))
//                        {
//                            PdfString Destination = AnnotationAction.GetAsString(PdfName.URI);
//                            string url1 = Destination.ToString();
//                        }
//                    }
//                }
//                res = qObj.SaveValidateResults(rObj);
//                rObj.CHECK_END_TIME = DateTime.Now;
//                return res;
//            }
//            //}

//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        //4.PDF Digital signature verification
//        public void VerifyPdfSignature(RegOpsQC rObj, string path, string destPath)
//        {

//            try
//            {
//                string res = string.Empty;
//                sourcePath = path + "//" + rObj.File_Name;
//                //destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                rObj.CHECK_START_TIME = DateTime.Now;
//                PdfReader reader = new PdfReader(sourcePath);
//                AcroFields af = reader.AcroFields;
//                var names = af.GetSignatureNames();

//                if (names.Count == 0)
//                {
//                    // throw new InvalidOperationException("No Signature present in pdf file.");
//                    rObj.QC_Result = "Passed";
//                    rObj.Comments = "No signature present in the given document.";
//                }

//                foreach (string name in names)
//                {
//                    //if (!af.SignatureCoversWholeDocument(name))
//                    //{
//                    //    throw new InvalidOperationException(string.Format("The signature: {0} does not covers the whole document.", name));
//                    //}

//                    PdfPKCS7 pk = af.VerifySignature(name);
//                    var cal = pk.SignDate;
//                    var pkc = pk.Certificates;

//                    if (!pk.Verify() && !pk.VerifyTimestampImprint())
//                    {
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "The signature and its timestamp could not be verified.";
//                        // throw new InvalidOperationException("The signature and its timestamp could not be verified.");
//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Passed";
//                        rObj.Comments = "The signature and its timestamp are verified.";
//                    }

//                    //X509Certificate2[] fails = CertificateVerification.VerifyCertificates(pkc, new X509Certificate[10], null, cal);
//                    //if (fails != null)
//                    //{
//                    //    throw new InvalidOperationException("The file is not signed using the specified key - pair.");
//                    //}
//                }
//                reader.Close();
//                rObj.CHECK_END_TIME = DateTime.Now;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                rObj.Job_Status = "Failed";
//                rObj.QC_Result = "Error";
//                rObj.Comments = "Technical error: " + ex.Message;                                
//            }
//        }


//    }
//}