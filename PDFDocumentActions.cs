using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Pdf;
using System.Configuration;
using CMCai.Models;
using Aspose.Pdf.Text;
using Aspose.Pdf.Facades;
//using Aspose.Pdf.Forms;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Devices;
//using GdPicture14;
using Bytescout.PDFExtractor;
using System.Collections;
//using GdPicture14;

namespace CMCai.Actions
{
    public class PDFDocumentActions
    {
        string sourcePath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString();

        string sourcePath = string.Empty;
        string destPath = string.Empty;
        int bookmarksflag = 0;
        Guid HOCRGUID;

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

        /// <summary>
        /// File name Length
        /// </summary>
        /// <param name="rObj"></param>
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
        /// <summary>
        /// No Security (password)-check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void Check_PDFFile_OpenPasswordProtection(RegOpsQC rObj, string path)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                PdfFileInfo fileInfo = new PdfFileInfo(sourcePath);
                string openprivilege = fileInfo.HasOpenPassword.ToString();
                if (openprivilege.ToLower() == "true")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document is password protected.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no 'Password Security' method for this file.";
                }
                fileInfo.Dispose();
                fileInfo.Close();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);

            }
        }
        /// <summary>
        /// No Edit Security-check & Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void Check_PDFFile_PasswordProtection(RegOpsQC rObj, string path, Document document)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);
                PdfFileInfo fileInfo = new PdfFileInfo(sourcePath);
                string privilege = fileInfo.HasEditPassword.ToString();
                if (privilege.ToLower() == "true" && rObj.Check_Type == 1)
                {
                    PdfFileSecurity fileSecurity = new PdfFileSecurity(document);
                    fileSecurity.DecryptFile("");
                    //document.Save(sourcePath);
                    //rObj.QC_Result = "Fixed";
                    rObj.QC_Result = "Failed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "No Edit Security is fixed";
                }
                else if (privilege.ToLower() == "true" && rObj.Check_Type == 0)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document has edit security";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no 'Edit Security' method for this file.";
                }
                fileInfo.Dispose();                                
                //document.Dispose();

                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);

            }
        }

        /// <summary>
        /// PDF Digital signature verification - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void VerifyPdfSignature(RegOpsQC rObj, string path , Document document)
        {
            try
            {
                string res = string.Empty;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
               
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
        /// Do OCR - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void EnableOCR(RegOpsQC rObj, string path, string destPath)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                MemoryStream ms = new MemoryStream();
                Document pdfdoc = new Document(sourcePath);
                PdfExtractor extractor = new PdfExtractor();
                extractor.BindPdf(sourcePath);
                Document doc = new Document(sourcePath);
                bool containsText = false;
                bool containsImage = false;
                extractor.ExtractImage();
                if (extractor.HasNextImage())
                    containsImage = true;

                if (containsImage == false)
                {
                    foreach (var page in pdfdoc.Pages)
                    {
                        extractor.StartPage = page.Number;
                        extractor.EndPage = page.Number;
                        extractor.ExtractText();
                        extractor.GetText(ms);
                        if (ms.Length <= 1)
                        {
                            containsText = false;
                            break;
                        }
                        else
                        {
                            containsText = true;
                            ms.SetLength(0);
                        }
                    }
                }
                if (containsText == true && containsImage == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "OCR is enabled";
                }
                else if (containsText == false || containsImage == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "OCR is disabled";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
                ms.Dispose();
            }
            catch (Exception ex)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
            }
        }

        /// <summary>
        /// Do OCR - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        //public void EnableOCRFix(RegOpsQC rObj, string path, string destPath)
        //{
        //    sourcePath = path + "//" + rObj.File_Name;
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    int flag = 0;
        //    string pageNumbers = string.Empty;
        //    try
        //    {
        //        HOCRGUID = Guid.NewGuid();
        //        Document document = new Document(sourcePath);
        //        Document pdfOutDcoument = new Document();
        //        string Tesseractdir = ConfigurationManager.AppSettings["TesseractPath"];
        //        string TesseractWorkingFolder = ConfigurationManager.AppSettings["TesseractWorkingFolder"];
        //        int pageCount = document.Pages.Count;
        //        for (int i = 1; i <= pageCount; i++)
        //        {
        //            try
        //            {
        //                using (FileStream imageStream = new FileStream(TesseractWorkingFolder + HOCRGUID + ".jpg", FileMode.Create))
        //                {
        //                    // Create JPEG device with specified attributes
        //                    // Width, Height, Resolution, Quality
        //                    // Quality [0-100], 100 is Maximum
        //                    // Create Resolution object
        //                    Resolution resolution = new Resolution(int.Parse(ConfigurationManager.AppSettings["TesseractResolutionWidth"]));

        //                    // JpegDevice jpegDevice = new JpegDevice(500, 700, resolution, 100);
        //                    JpegDevice jpegDevice = new JpegDevice(resolution, int.Parse(ConfigurationManager.AppSettings["TesseractResolutionHeight"]));

        //                    // Convert a particular page and save the image to stream
        //                    jpegDevice.Process(document.Pages[i], imageStream);

        //                    // Close stream
        //                    imageStream.Close();
        //                }
        //                ProcessStartInfo info = new ProcessStartInfo(Tesseractdir + "tesseract");
        //                info.WindowStyle = ProcessWindowStyle.Hidden;
        //                info.Arguments = "\"" + TesseractWorkingFolder + HOCRGUID + ".jpg\" \"" + TesseractWorkingFolder + HOCRGUID + "\" pdf";
        //                Process p = new Process();
        //                p.StartInfo = info;
        //                p.Start();
        //                p.WaitForExit();
        //                Document TempDocument = new Document(TesseractWorkingFolder + HOCRGUID + ".pdf");
        //                pdfOutDcoument.Pages.Add(TempDocument.Pages);
        //            }
        //            catch (Exception ex)
        //            {
        //                flag = 1;
        //                pageNumbers = pageNumbers + i + ", ";
        //                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
        //            }
        //        }
        //        pdfOutDcoument.Save(sourcePath);
        //        if (flag == 1)
        //        {
        //            rObj.Job_Status = "Error";
        //            rObj.QC_Result = "Failed";
        //            rObj.Comments = "Error in converting the page(s): " + pageNumbers.Trim().TrimEnd(',');
        //        }
        //        else
        //        {
        //            rObj.QC_Result = "Fixed";
        //            rObj.Comments = "OCR is enabled";
        //        }
        //        if (File.Exists(TesseractWorkingFolder + HOCRGUID + ".jpg"))
        //            File.Delete(TesseractWorkingFolder + HOCRGUID + ".jpg");
        //        if (File.Exists(TesseractWorkingFolder + HOCRGUID + ".pdf"))
        //            File.Delete(TesseractWorkingFolder + HOCRGUID + ".pdf");
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

        ///// <summary>
        ///// Do OCR using GD Picture - fix
        ///// </summary>
        ///// <param name="rObj"></param>
        ///// <param name="path"></param>
        ///// <param name="destPath"></param>
        //public void EnableOCRFixGDPicture(RegOpsQC rObj, string path, string destPath)
        //{
        //    sourcePath = path + "//" + rObj.File_Name;
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    try
        //    {
        //        Document document = new Document(sourcePath);

        //        // properties to blank
        //        DocumentInfo docInfo = document.Info;
        //        string Title = string.Empty;
        //        string Subject = string.Empty;
        //        string Author = string.Empty;
        //        string Keywords = string.Empty;
        //        foreach (var d in docInfo)
        //        {
        //            if (d.Key == "Title")
        //                Title = d.Value;
        //            if (d.Key == "Subject")
        //                Subject = d.Value;
        //            if (d.Key == "Author")
        //                Author = d.Value;
        //            if (d.Key == "Keywords")
        //                Keywords = d.Value;
        //        }

        //        HOCRGUID = Guid.NewGuid();
        //        //We assume that GdPicture has been correctly installed and unlocked.
        //        GdPicturePDF oGdPicturePDF = new GdPicturePDF();
        //        //Loading an input document.     
        //        String TempFileName = path + "\\" + HOCRGUID + ".pdf";
        //        string GDRPath = ConfigurationManager.AppSettings["GDOCRResources"];
        //        GdPictureStatus status = oGdPicturePDF.LoadFromFile(sourcePath, false);
        //        //Checking if loading has been successful.
        //        int flag = 0;
        //        string pageNumbers = string.Empty;
        //        if (status == GdPictureStatus.OK)
        //        {
        //            int pageCount = oGdPicturePDF.GetPageCount();
        //            //Loop through pages.
        //            for (int i = 1; i <= pageCount; i++)
        //            {
        //                //Selecting a page.
        //                oGdPicturePDF.SelectPage(i);
        //                if (oGdPicturePDF.OcrPage("eng", GDRPath, "", 200) != GdPictureStatus.OK)
        //                {
        //                    flag = 1;
        //                    pageNumbers = pageNumbers + i + ", ";
        //                    rObj.Job_Status = "Error";
        //                    rObj.QC_Result = "Error";
        //                    rObj.Comments = "Technical error: " + oGdPicturePDF.GetStat().ToString();
        //                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + oGdPicturePDF.GetStat().ToString());
        //                }
        //            }
        //            //Saving to a different file.
        //            status = oGdPicturePDF.SaveToFile(TempFileName, true);
        //            if (flag == 1)
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Error in converting the page(s): " + pageNumbers.Trim().TrimEnd(',');
        //            }
        //            else if (flag != 1 && status == GdPictureStatus.OK)
        //            {
        //                rObj.QC_Result = "Fixed";
        //                rObj.Comments = "OCR is enabled";
        //            }
        //            else if(status != GdPictureStatus.OK)
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Error in converting the document";
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
        //        Document document1 = new Document(TempFileName);

        //        // properties to blank
        //        DocumentInfo docInfo1 = document1.Info;
        //        document1.RemoveMetadata();
        //        docInfo1.Title = Title;
        //        docInfo1.Author = Author;
        //        docInfo1.Subject = Subject;
        //        docInfo1.Keywords = Keywords;
        //        //docInfo.Remove("Author");
        //        //docInfo.Remove("Subject");
        //        //docInfo.Remove("Keywords");
        //        document1.Save(sourcePath);
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

 

        public void EnableBytesCoutOCRFix(RegOpsQC rObj, string path, Document document)
        {
            sourcePath = path + "//" + rObj.File_Name;
            document.Save(sourcePath);          
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                PdfFileInfo fileInfo = new PdfFileInfo(document);
                string privilege = fileInfo.HasEditPassword.ToString();
                if (privilege.ToLower() == "true")
                {
                   // fileInfo.Dispose();
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "OCR cannot be enabled because document has edit password";
                }
                else
                {
                  //  fileInfo.Dispose();

                    Guid guid = Guid.NewGuid();
                    string BytescoutLicenceName = ConfigurationManager.AppSettings["BytescoutLicenceName"];
                    string BytescoutLicenceKey = ConfigurationManager.AppSettings["BytescoutLicenceKey"];
                    string Bytescoutdir = ConfigurationManager.AppSettings["BytescoutOCRLanguageDataFolder"];
                    Decimal BytescoutOCRResolution = Convert.ToDecimal(ConfigurationManager.AppSettings["BytescoutOCRResolution"]);

                    using (var searchablePDFMaker = new SearchablePDFMaker(BytescoutLicenceName, BytescoutLicenceKey))
                    {
                        // Load sample PDF document
                        searchablePDFMaker.LoadDocumentFromFile(sourcePath);

                        searchablePDFMaker.OCRMaximizeCPUUtilization = true;
                        //searchablePDFMaker.OCRImagePreprocessingFilters.AddDeskew();
                        //searchablePDFMaker.OCRImagePreprocessingFilters.AddDilate();
                        searchablePDFMaker.OCRDetectPageRotation = true;
                        // Set the location of OCR language data files
                        searchablePDFMaker.OCRLanguageDataFolder = Bytescoutdir;

                        // Set OCR language
                        searchablePDFMaker.OCRLanguage = "eng"; // "eng" for english, "deu" for German, "fra" for French, "spa" for Spanish etc - according to files in "ocrdata" folder
                                                                // Find more language files at https://github.com/bytescout/ocrdata

                        // Set PDF document rendering resolution
                        searchablePDFMaker.OCRResolution = (float)BytescoutOCRResolution;

                        // Save extracted text to file
                        searchablePDFMaker.MakePDFSearchable(path + "//" + guid + rObj.File_Name);
                    }
                    File.Copy(path + "//" + guid + rObj.File_Name, sourcePath, true);
                    
                    if (File.Exists(path + "//" + guid + rObj.File_Name))
                        File.Delete(path + "//" + guid + rObj.File_Name);

                    // As suggested and issues #2860, we wil capture only Fix as yes and QC_RESULT will be blank
                  //  rObj.QC_Result = "Passed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "OCR is Enabled";
                    rObj.FIX_END_TIME = DateTime.Now;
                }

            }
            catch (Exception ex)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
            }
        }
        /// PDF version 
        public void PDFVersionCheck(RegOpsQC rObj, List<RegOpsQC> chLst, Document document)
        {
            string res = string.Empty;
            try
            {                
                rObj.CHECK_START_TIME = DateTime.Now;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                List<string> VersionLst = new List<string>();
                bool isvalid = false;
                string fixedversion = string.Empty;
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;
                    rObj.CHECK_START_TIME = DateTime.Now;                    
                }
                for (int i = 0; i < chLst.Count; i++)
                {
                    if (chLst[i].Check_Name.ToString() == "Valid Version(s)")
                    {
                        string[] VersionAry = new string[] { };

                        if (chLst[i].Check_Parameter != null)
                        {
                            VersionAry = chLst[i].Check_Parameter.Split(',');
                            for (int a = 0; a < VersionAry.Length; a++)
                            {
                                string exceptionfont = VersionAry[a].Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");
                                VersionLst.Add(exceptionfont);
                            }
                        }
                    }
                    else if (chLst[i].Check_Name.ToString() == "Fix to Version" && chLst[i].Check_Type == 1)
                    {
                        fixedversion = chLst[i].Check_Parameter;
                    }
                }
                string version = document.Version;          
                foreach (string s in VersionLst)
                {
                    if (s == version)
                    {
                        isvalid = true;
                        break;
                    }

                }
                if (isvalid || fixedversion == version)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "PDF Version is : " + version + ".";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "PDF Version is not in selected version(s)";

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
        /// PDF version Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void Check_PDFVersionFix(RegOpsQC rObj, List<RegOpsQC> chLst, Document document)
        {
            string res = string.Empty;
            try
            {
                bool flag = false;

                rObj.FIX_START_TIME = DateTime.Now;
                //sourcePath = path + "//" + rObj.File_Name;
                // Document document = new Document(sourcePath);
                string version = document.Version;
                string fixedversion = string.Empty;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;
                    rObj.CHECK_START_TIME = DateTime.Now;                    
                }
                for (int i = 0; i < chLst.Count; i++)
                {
                    if (chLst[i].Check_Name.ToString() == "Fix to Version" && chLst[i].Check_Type == 1)
                    {
                        fixedversion = chLst[i].Check_Parameter;
                    }
                }
                if (document.Pages.Count != 0)
                {
                    if (fixedversion != null && fixedversion != "")
                    {
                        if (fixedversion == "1.4")
                        {
                            document.Convert(new MemoryStream(), PdfFormat.v_1_4, ConvertErrorAction.None);
                            flag = true;
                        }
                        else if (fixedversion == "1.5")
                        {
                            document.Convert(new MemoryStream(), PdfFormat.v_1_5, ConvertErrorAction.None);
                            flag = true;
                        }
                        else if (fixedversion == "1.6")
                        {
                            document.Convert(new MemoryStream(), PdfFormat.v_1_6, ConvertErrorAction.None);
                            flag = true;
                        }
                        else if (fixedversion == "1.7")
                        {
                            document.Convert(new MemoryStream(), PdfFormat.v_1_7, ConvertErrorAction.None);
                            flag = true;
                        }

                    }
                    if (flag == true)
                    {
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = rObj.Comments + ". Fixed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }
                //document.Save(sourcePath);
                //document.Dispose();
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
        /// PDF version 1.4 to 1.7
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void Check_PDFVersion(RegOpsQC rObj, string path, Document document)
        {
            string res = string.Empty;
            try
            {
                //sourcePath = path;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);
                string version = document.Version;
                if (version == "1.3" || version == "1.4" || version == "1.5" || version == "1.6" || version =="1.7")
                {
                    rObj.QC_Result = "Passed";
                   // rObj.Comments = "PDF Version is : " + version + "";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "PDF Version is not in selected version(s)";
                }
                //document.Dispose();
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
        /// Remove Blank pages - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void RemoveBlankPages(RegOpsQC rObj, string path,Document pdfdoc)
        {
            string Pagenumber = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                bool flag = false;
                //Document pdfdoc = new Document(sourcePath);
                List<int> pgnumLst = new List<int>();
                if (pdfdoc.Pages.Count != 0)
                {
                    //PdfExtractor extractor = new PdfExtractor();
                    //extractor.BindPdf(sourcePath);
                    //foreach (var page in pdfdoc.Pages)
                    //{
                    //    using (page)
                    //    {
                    //        extractor.StartPage = page.Number;
                    //        extractor.EndPage = page.Number;
                    //        extractor.ExtractImage();
                    //        if (!(extractor.HasNextImage()))
                    //        {
                    //            if (IsBlankPage(page, pdfdoc, page.Number))
                    //            {
                    //                count = count + 1;
                    //            }
                    //        }
                    //    }
                    //}

                    // Testing purpose
                    //for (int i = 1; i <= pdfdoc.Pages.Count(); i++)
                    //{
                    //    using (Page page = pdfdoc.Pages[i])
                    //    {
                    //        if (page.IsBlank(0.01))
                    //        {
                    //            flag = true;
                    //            pgnumLst.Add(page.Number);
                    //        }
                    //    }
                    //}
                    //End of testing code
                    foreach (var page in pdfdoc.Pages)
                    {
                        if (page.IsBlank(0.01))
                        {
                            flag = true;
                            pgnumLst.Add(page.Number);
                        }
                        page.FreeMemory();
                    }

                    if (flag == false)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No blank pages exists in given document.";
                    }
                    else
                    {
                        List<int> lst2 = pgnumLst.Distinct().ToList();
                        if (lst2.Count > 0)
                        {
                            lst2.Sort();
                            Pagenumber = string.Join(", ", lst2.ToArray());
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Blank Page are found in: " + Pagenumber;
                            rObj.CommentsWOPageNum = "Blank page found";
                            rObj.PageNumbersLst = lst2;
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }                
                //pdfdoc.Dispose();
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
        /// Remove Blank pages - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void RemoveBlankPagesFix(RegOpsQC rObj, string path, Document pdfdoc)
        {
            string res = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                bool flag = false;
                //Document pdfdoc = new Document(sourcePath);
                List<int> pgnumLst = new List<int>();
                if (pdfdoc.Pages.Count != 0)
                {
                    foreach (var page in pdfdoc.Pages)
                    {
                        if (page.IsBlank(0.01))
                        {
                            flag = true;
                            pgnumLst.Add(page.Number);
                        }
                        page.FreeMemory();
                    }
                    if (flag == true)
                    {
                        pdfdoc.Pages.Delete(pgnumLst.ToArray());
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = rObj.Comments + ". These are removed";
                        rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }
                //pdfdoc.Save(sourcePath);
                //pdfdoc.Dispose();
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

        public static bool IsBlankPage(Aspose.Pdf.Page page, Document pdfdoc, int i)
        {
            if ((page.Contents.Count == 0 && page.Annotations.Count == 0) && HasOnlyWhiteImages(page))
            {
                return true;
            }
            else
            {
                // commented below working code due to OCR issue.

                //TextAbsorber textAbsorber = new TextAbsorber();
                //pdfdoc.Pages[i].Accept(textAbsorber);
                //string extractedText = textAbsorber.Text;
                //if (extractedText.Replace("\n", "").Replace("\r", "").Trim() == "")
                //    return true;
                //else
                //    return false;
                return false;
            }
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

        /// <summary>
        /// Properties fields should be blank - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="checkString"></param>
        /// <param name="destPath"></param>
        /// <param name="checkType"></param>
        public void PDFFile_Properties(RegOpsQC rObj, string path, string checkString, double checkType,Document document)
        {
            string res = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);

                // properties to blank
                DocumentInfo docInfo = document.Info;
                string result = string.Empty;
                foreach (var d in docInfo)
                {
                    if ((d.Key == "Title" || d.Key == "Subject" || d.Key == "Author" || d.Key == "Keywords") && d.Value != "")
                        result = result + " , " + d.Key + ": " + d.Value;
                }                
                if (result == "")
                {
                    Regex rx_Prop = new Regex(@"dc:title|dc:Subject|dc:Author|dc:Keywords",RegexOptions.IgnoreCase);
                    foreach (var d in document.Metadata)
                    {
                        if (rx_Prop.IsMatch(d.Key) && d.Value.ToString() != "")
                            result = result + " , " + d.Key + ": " + d.Value;
                    }
                }
                if (result != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Default properties are not empty";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Default properties are empty";
                }                
                //document.Dispose();
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
        /// Properties fields should be blank - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="checkString"></param>
        /// <param name="destPath"></param>
        /// <param name="checkType"></param>
        public void PDFFile_PropertiesFix(RegOpsQC rObj, string path, string checkString, double checkType,Document document)
        {
            string res = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);

                // properties to blank
                DocumentInfo docInfo = document.Info;
                string result = string.Empty;

                document.RemoveMetadata();
                docInfo.Remove("Title");
                docInfo.Remove("Author");
                docInfo.Remove("Subject");
                docInfo.Remove("Keywords");

                //rObj.QC_Result = "Fixed";
                rObj.Is_Fixed = 1;
                rObj.Comments = "Default properties set to blank";
                //document.Save(sourcePath);
                //document.Dispose();
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
        /// Save as Optimized PDF - Direct fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void SaveAsOptimized(RegOpsQC rObj, string path,Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                pdfDocument.Save(sourcePath);
                Document pdfDocument1 = new Document(sourcePath);

                string ext = Path.GetExtension(rObj.File_Name);
                Guid g = Guid.NewGuid();
                try
                {
                    pdfDocument1.Optimize();
                    pdfDocument1.OptimizeSize = true;
                    pdfDocument1.OptimizeResources();
                    pdfDocument1.Save(path + "//" + rObj.File_Name.Replace(ext, g + ext));
                    Document pdfDocument2 = new Document(path + "//" + rObj.File_Name.Replace(ext, g + ext));
                    pdfDocument2.Save(sourcePath);
                    pdfDocument1.Dispose();
                    pdfDocument2.Dispose();

                    if (File.Exists(path + "//" + rObj.File_Name.Replace(ext, g + ext)))
                        File.Delete(path + "//" + rObj.File_Name.Replace(ext, g + ext));

                    // As suggested and issues #2860, we wil capture only Fix as yes and QC_RESULT will be blank
                    //  rObj.QC_Result = "Passed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "Document saved as optimized pdf";
                }
                catch
                {
                    try
                    {
                        Document TemppdfDocument = new Document(sourcePath);
                        TemppdfDocument.OptimizeSize = true;
                        TemppdfDocument.OptimizeResources();
                        TemppdfDocument.Save(path + "//" + rObj.File_Name.Replace(ext, g + ext));
                        Document Temppdf = new Document(path + "//" + rObj.File_Name.Replace(ext, g + ext));
                        Temppdf.Save(sourcePath);
                        Temppdf.Dispose();
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Unable to optimize PDF";
                        if (File.Exists(path + "//" + rObj.File_Name.Replace(ext, g + ext)))
                            File.Delete(path + "//" + rObj.File_Name.Replace(ext, g + ext));
                    }
                    catch
                    {
                        Document TemppdfDocument2 = new Document(sourcePath);
                        TemppdfDocument2.OptimizeSize = true;
                        TemppdfDocument2.Save(sourcePath);
                        TemppdfDocument2.Dispose();
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Unable to optimize PDF";
                        if (File.Exists(path + "//" + rObj.File_Name.Replace(ext, g + ext)))
                            File.Delete(path + "//" + rObj.File_Name.Replace(ext, g + ext));
                    }
                }

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
        /// Remove the original Pedigree - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void RemoveOriginalPedigree(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.Comments = string.Empty;
            bool isValid = true;
            sourcePath = path + "//" + rObj.File_Name;
            string[] strPedigree = null;
            bool isDoublePedigree = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            string pageNumbers = string.Empty;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                //pdfDocument = new Document(sourcePath);
                for (int j = 1; j <= pdfDocument.Pages.Count; j++)
                {
                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                    Page page = pdfDocument.Pages[j];
                    page.Accept(textFragmentAbsorber);
                    // Get the extracted text fragments
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                    TextFragment textFragmentPed;
                    for (int i = 1; i <= textFragmentCollection.Count(); i++)
                    {
                        textFragmentPed = textFragmentCollection[i];
                        bool hexaValue = false;
                        bool dateValue = false;
                        strPedigree = null;
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
                                    rObj.Comments = "Pedigree(s) exists in: ";
                                }
                                if (!rObj.Comments.Contains(textFragmentPed.Page.Number.ToString() + ", "))
                                {
                                    rObj.Comments = rObj.Comments + textFragmentPed.Page.Number + ", ";
                                    pageNumbers = pageNumbers + textFragmentPed.Page.Number + ", ";
                                }                                    
                                isValid = false;
                                rObj.QC_Result = "Failed";
                            }
                        }
                    }
                    page.FreeMemory();
                }

                if (rObj.Comments != "")
                {
                    rObj.Comments = rObj.Comments.Trim().TrimEnd(',');
                    if(pageNumbers != "")
                    rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    rObj.CommentsWOPageNum = "Pedigree exists";
                }
                if (isDoublePedigree)
                {
                    rObj.Comments = "Double pedigree found";
                    rObj.QC_Result = "Failed";
                }
                else if (isValid == true)
                {
                    //rObj.Comments = "No Pedigree exists in the document(s)";
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

        /// <summary>
        /// Remove the original Pedigree - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void RemoveOriginalPedigreeFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
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
                                if (Regex.IsMatch(strPedigree[k], @"(\d{2}\s?\-[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})", RegexOptions.IgnoreCase))
                                {
                                    //Match m = Regex.Match(strPedigree[k], @"(\d{2}\s?\-\s?[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})", RegexOptions.IgnoreCase);
                                    dateValue = true;
                                }
                            }
                            if (hexaValue && dateValue)
                            {
                                textFragmentCollection.Remove(textFragmentPed);
                                textFragmentPed.Text = "";
                                textFragmentPed = null;
                                //rObj.QC_Result = "Fixed";
                                rObj.Is_Fixed = 1;
                                j--;
                            }
                        }

                    }
                }
                if (rObj.Is_Fixed == 1)
                {
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                    rObj.Comments = rObj.Comments.Trim().TrimEnd(',') + ". Fixed";
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

        //public void LinkAttributor(RegOpsQC rObj, string path, List<RegOpsQC> chLst)
        //{
        //    string res = string.Empty;
        //    sourcePath = path + "//" + rObj.File_Name;
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    string textcolor = string.Empty;
        //    string checkname = string.Empty;
        //    string Linebordercolor = string.Empty;
        //    try
        //    {
        //        // to get sub check list
        //        chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
        //        rObj.CHECK_START_TIME = DateTime.Now;
        //        string FailedFlag = string.Empty;
        //        string PassedFlag = string.Empty;
        //        string FailedFlag1 = string.Empty;
        //        string PassedFlag1 = string.Empty;
        //        int pgnum = 0;
        //        string bothExists = "";
        //        for (int i = 0; i < chLst.Count; i++)
        //        {
        //            chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
        //            chLst[i].JID = rObj.JID;
        //            chLst[i].Job_ID = rObj.Job_ID;
        //            chLst[i].Folder_Name = rObj.Folder_Name;
        //            chLst[i].File_Name = rObj.File_Name;
        //            chLst[i].Created_ID = rObj.Created_ID;

        //            if (chLst[i].Check_Name.ToString() == "Text Color")
        //            {
        //                textcolor = chLst[i].Check_Parameter.ToString();
        //            }

        //            if (chLst[i].Check_Name.ToString() == "Link Border Color")
        //            {
        //                Linebordercolor = chLst[i].Check_Parameter.ToString();
        //            }

        //        }

        //        Document document = new Document(sourcePath);
        //        if (document.Pages.Count != 0)
        //        {

        //            var editor = new PdfContentEditor(document);
        //            string pageNumbers = "";
        //            List<PageNumberReport> pglst = new List<PageNumberReport>();

        //            // Linebordercolor = chLst[1].Check_Parameter.ToString();
        //            foreach (Aspose.Pdf.Page page in document.Pages)
        //            {
        //                PageNumberReport pgObj = new PageNumberReport();
        //                FailedFlag1 = string.Empty;
        //                PassedFlag1 = string.Empty;
        //                pgnum = 0;
        //                // Get the link annotations from particular page
        //                AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
        //                page.Accept(selector);
        //                // Create list holding all the links
        //                IList<Annotation> list = selector.Selected;
        //                // Iterate through invidiaul item inside list                   
        //                foreach (LinkAnnotation a in list)
        //                {
        //                    string URL = string.Empty;
        //                    string URL1 = string.Empty;
        //                    string URL2 = string.Empty;
        //                    IAppointment dest = a.Destination;
        //                    //Checking whether link has a destination/action  or not.
        //                    if (a.Action != null || dest != null)
        //                    {

        //                        #region action for Aspose.Pdf.Annotations.GoToAction                                
        //                        if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
        //                        {

        //                            URL2 = ((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString();
        //                            if (URL2 != "")
        //                            {
        //                                if (dest != null)
        //                                {
        //                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
        //                                    Aspose.Pdf.Rectangle rect = a.Rect;
        //                                    ta.TextSearchOptions = new TextSearchOptions(rect);
        //                                    ta.Visit(page);
        //                                    if (ta.TextFragments.Count > 0)
        //                                    {
        //                                        foreach (TextFragment tf in ta.TextFragments)
        //                                        {
        //                                            if (tf.TextState.Invisible == false)
        //                                            {
        //                                                Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
        //                                                string colortext = color1.ToString();
        //                                                if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
        //                                                {
        //                                                    FailedFlag = "Failed";
        //                                                    FailedFlag1 = "Failed";
        //                                                    pgnum = page.Number;
        //                                                    if (pageNumbers == "")
        //                                                    {
        //                                                        pageNumbers = page.Number.ToString() + ", ";
        //                                                    }
        //                                                    else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                                }
        //                                                else
        //                                                {
        //                                                    int border = a.Border.Width;
        //                                                    Aspose.Pdf.Color lncolor = a.Color;
        //                                                    if (border != 0)
        //                                                    {
        //                                                        Border b = new Border(a);
        //                                                        b.Width = 0;
        //                                                        b.Style = BorderStyle.Solid;
        //                                                        FailedFlag = "Failed";
        //                                                        FailedFlag1 = "Failed";
        //                                                        pgnum = page.Number;
        //                                                        bothExists = FailedFlag;
        //                                                        if (pageNumbers == "")
        //                                                        {
        //                                                            pageNumbers = page.Number.ToString() + ", ";
        //                                                        }
        //                                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                                    }
        //                                                    else
        //                                                    {
        //                                                        PassedFlag = "Passed";
        //                                                        PassedFlag1 = "Passed";
        //                                                    }

        //                                                }
        //                                            }
        //                                            else
        //                                            {
        //                                                if (Linebordercolor != "")
        //                                                {
        //                                                    Aspose.Pdf.Color color = GetColor(Linebordercolor);
        //                                                    int border = a.Border.Width;
        //                                                    Aspose.Pdf.Color lncolor = a.Color;
        //                                                    if (color.ToString().ToUpper() != lncolor.ToString().ToUpper())
        //                                                    {
        //                                                        FailedFlag = "Failed";
        //                                                        FailedFlag1 = "Failed";
        //                                                        pgnum = page.Number;
        //                                                        if (pageNumbers == "")
        //                                                        {
        //                                                            pageNumbers = page.Number.ToString() + ", ";
        //                                                        }
        //                                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                                    }
        //                                                    else
        //                                                    {
        //                                                        border = a.Border.Width;
        //                                                        lncolor = a.Color;
        //                                                        if (border != 0)
        //                                                        {
        //                                                            Border b = new Border(a);
        //                                                            b.Width = 0;
        //                                                            b.Style = BorderStyle.Solid;
        //                                                            FailedFlag = "Failed";
        //                                                            FailedFlag1 = "Failed";
        //                                                        }
        //                                                        else
        //                                                        {
        //                                                            PassedFlag = "Passed";
        //                                                            PassedFlag1 = "Passed";
        //                                                        }
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        if (Linebordercolor != "")
        //                                        {
        //                                            Aspose.Pdf.Color color = GetColor(Linebordercolor);
        //                                            int border = a.Border.Width;
        //                                            Aspose.Pdf.Color lncolor = a.Color;
        //                                            if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
        //                                            {
        //                                                FailedFlag = "Failed";
        //                                                FailedFlag1 = "Failed";
        //                                                pgnum = page.Number;
        //                                                if (pageNumbers == "")
        //                                                {
        //                                                    pageNumbers = page.Number.ToString() + ", ";
        //                                                }
        //                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                            }
        //                                            else
        //                                            {
        //                                                border = a.Border.Width;
        //                                                lncolor = a.Color;
        //                                                if (border != 0)
        //                                                {
        //                                                    Border b = new Border(a);
        //                                                    b.Width = 0;
        //                                                    b.Style = BorderStyle.Solid;
        //                                                    FailedFlag = "Failed";
        //                                                    FailedFlag1 = "Failed";
        //                                                }
        //                                                else
        //                                                {
        //                                                    PassedFlag = "Passed";
        //                                                    PassedFlag1 = "Passed";
        //                                                }
        //                                            }
        //                                        }
        //                                    }


        //                                }
        //                                else
        //                                {
        //                                    TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
        //                                    Aspose.Pdf.Rectangle rect = a.Rect;
        //                                    ta1.TextSearchOptions = new TextSearchOptions(rect);
        //                                    ta1.Visit(page);
        //                                    if (ta1.TextFragments.Count > 0)
        //                                    {
        //                                        foreach (TextFragment tf in ta1.TextFragments)
        //                                        {
        //                                            string txt = tf.Text;
        //                                            if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
        //                                            {
        //                                                if (tf.TextState.Invisible == false)
        //                                                {
        //                                                    Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
        //                                                    string colortext = color1.ToString();
        //                                                    if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
        //                                                    {
        //                                                        FailedFlag = "Failed";
        //                                                        FailedFlag1 = "Failed";
        //                                                        pgnum = page.Number;
        //                                                        if (pageNumbers == "")
        //                                                        {
        //                                                            pageNumbers = page.Number.ToString() + ", ";
        //                                                        }
        //                                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                                    }
        //                                                    else
        //                                                    {
        //                                                        int border = a.Border.Width;
        //                                                        Aspose.Pdf.Color lncolor = a.Color;
        //                                                        if (border != 0)
        //                                                        {
        //                                                            Border b = new Border(a);
        //                                                            b.Width = 0;
        //                                                            b.Style = BorderStyle.Solid;
        //                                                            FailedFlag = "Failed";
        //                                                            FailedFlag1 = "Failed";
        //                                                            pgnum = page.Number;
        //                                                            bothExists = FailedFlag;
        //                                                            if (pageNumbers == "")
        //                                                            {
        //                                                                pageNumbers = page.Number.ToString() + ", ";
        //                                                            }
        //                                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                                        }
        //                                s                        else
        //                                                        {
        //                                                            PassedFlag = "Passed";
        //                                                            PassedFlag1 = "Passed";
        //                                                        }
        //                                                    }
        //                                                }
        //                                                else
        //                                                {
        //                                                    if (Linebordercolor != "")
        //                                                    {
        //                                                        Aspose.Pdf.Color color = GetColor(Linebordercolor);
        //                                                        int border = a.Border.Width;
        //                                                        Aspose.Pdf.Color lncolor = a.Color;
        //                                                        if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
        //                                                        {
        //                                                            FailedFlag = "Failed";
        //                                                            FailedFlag1 = "Failed";
        //                                                            pgnum = page.Number;
        //                                                            if (pageNumbers == "")
        //                                                            {
        //                                                                pageNumbers = page.Number.ToString() + ", ";
        //                                                            }
        //                                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                                        }
        //                                                        else
        //                                                        {
        //                                                            border = a.Border.Width;
        //                                                            lncolor = a.Color;
        //                                                            if (border == 0)
        //                                                            {
        //                                                                Border b = new Border(a);
        //                                                                b.Width = 0;
        //                                                                b.Style = BorderStyle.Solid;
        //                                                                FailedFlag = "Failed";
        //                                                                FailedFlag1 = "Failed";
        //                                                            }
        //                                                            else
        //                                                            {
        //                                                                PassedFlag = "Passed";
        //                                                                PassedFlag1 = "Passed";
        //                                                            }
        //                                                        }
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        if (Linebordercolor != "")
        //                                        {
        //                                            Aspose.Pdf.Color color = GetColor(Linebordercolor);
        //                                            int border = a.Border.Width;
        //                                            Aspose.Pdf.Color lncolor = a.Color;
        //                                            if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
        //                                            {
        //                                                FailedFlag = "Failed";
        //                                                FailedFlag1 = "Failed";
        //                                                pgnum = page.Number;
        //                                                if (pageNumbers == "")
        //                                                {
        //                                                    pageNumbers = page.Number.ToString() + ", ";
        //                                                }
        //                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                            }
        //                                            else
        //                                            {
        //                                                border = a.Border.Width;
        //                                                lncolor = a.Color;
        //                                                if (border != 0)
        //                                                {
        //                                                    Border b = new Border(a);
        //                                                    b.Width = 0;
        //                                                    b.Style = BorderStyle.Solid;
        //                                                    FailedFlag = "Failed";
        //                                                    FailedFlag1 = "Failed";
        //                                                }
        //                                                else
        //                                                {
        //                                                    PassedFlag = "Passed";
        //                                                    PassedFlag1 = "Passed";
        //                                                }
        //                                            }
        //                                        }
        //                                    }

        //                                }
        //                            }
        //                        }
        //                        #endregion
        //                        #region action is null                                
        //                        if (dest != null)
        //                        {
        //                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
        //                            Aspose.Pdf.Rectangle rect = a.Rect;
        //                            ta.TextSearchOptions = new TextSearchOptions(rect);
        //                            ta.Visit(page);
        //                            if (ta.TextFragments.Count > 0)
        //                            {
        //                                foreach (TextFragment tf in ta.TextFragments)
        //                                {
        //                                    if (tf.TextState.Invisible == false)
        //                                    {
        //                                        Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
        //                                        string colortext = color1.ToString();
        //                                        if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
        //                                        {
        //                                            FailedFlag = "Failed";
        //                                            FailedFlag1 = "Failed";
        //                                            pgnum = page.Number;
        //                                            if (pageNumbers == "")
        //                                            {
        //                                                pageNumbers = page.Number.ToString() + ", ";
        //                                            }
        //                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                        }
        //                                        else
        //                                        {
        //                                            int border = a.Border.Width;
        //                                            Aspose.Pdf.Color lncolor = a.Color;
        //                                            if (border != 0)
        //                                            {
        //                                                Border b = new Border(a);
        //                                                b.Width = 0;
        //                                                b.Style = BorderStyle.Solid;
        //                                                FailedFlag = "Failed";
        //                                                FailedFlag1 = "Failed";
        //                                                pgnum = page.Number;
        //                                                bothExists = FailedFlag;
        //                                                if (pageNumbers == "")
        //                                                {
        //                                                    pageNumbers = page.Number.ToString() + ", ";
        //                                                }
        //                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                            }
        //                                            else
        //                                            {
        //                                                PassedFlag = "Passed";
        //                                                PassedFlag1 = "Passed";
        //                                            }
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        if (Linebordercolor != "")
        //                                        {
        //                                            Aspose.Pdf.Color color = GetColor(Linebordercolor);
        //                                            int border = a.Border.Width;
        //                                            Aspose.Pdf.Color lncolor = a.Color;
        //                                            if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
        //                                            {
        //                                                FailedFlag = "Failed";
        //                                                FailedFlag1 = "Failed";
        //                                                pgnum = page.Number;
        //                                                if (pageNumbers == "")
        //                                                {
        //                                                    pageNumbers = page.Number.ToString() + ", ";
        //                                                }
        //                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                            }
        //                                            else
        //                                            {
        //                                                border = a.Border.Width;
        //                                                lncolor = a.Color;
        //                                                if (border != 0)
        //                                                {
        //                                                    Border b = new Border(a);
        //                                                    b.Width = 0;
        //                                                    b.Style = BorderStyle.Solid;
        //                                                    FailedFlag = "Failed";
        //                                                    FailedFlag1 = "Failed";
        //                                                }
        //                                                else
        //                                                {
        //                                                    PassedFlag = "Passed";
        //                                                    PassedFlag1 = "Passed";
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                            else
        //                            {
        //                                if (Linebordercolor != "")
        //                                {
        //                                    Aspose.Pdf.Color color = GetColor(Linebordercolor);
        //                                    int border = a.Border.Width;
        //                                    Aspose.Pdf.Color lncolor = a.Color;
        //                                    if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
        //                                    {
        //                                        FailedFlag = "Failed";
        //                                        FailedFlag1 = "Failed";
        //                                        pgnum = page.Number;
        //                                        if (pageNumbers == "")
        //                                        {
        //                                            pageNumbers = page.Number.ToString() + ", ";
        //                                        }
        //                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                    }
        //                                    else
        //                                    {
        //                                        border = a.Border.Width;
        //                                        lncolor = a.Color;
        //                                        if (border != 0)
        //                                        {
        //                                            Border b = new Border(a);
        //                                            b.Width = 0;
        //                                            b.Style = BorderStyle.Solid;
        //                                            FailedFlag = "Failed";
        //                                            FailedFlag1 = "Failed";
        //                                        }
        //                                        else
        //                                        {
        //                                            PassedFlag = "Passed";
        //                                            PassedFlag1 = "Passed";
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                        }

        //                        #endregion
        //                    }
        //                }
        //                if (FailedFlag1 != "" && PassedFlag1 != "" && pgnum == 0)
        //                {
        //                    pgObj.PageNumber = page.Number;
        //                    pgObj.Comments = "Link text color matched but border also existed for the links.";
        //                    pglst.Add(pgObj);
        //                }
        //                else if (FailedFlag1 != "")
        //                {
        //                    pgObj.PageNumber = page.Number;
        //                    rObj.Comments = "No Link border color or Text color.";
        //                    pglst.Add(pgObj);
        //                }
        //                else if (FailedFlag1 == "" && PassedFlag1 != "")
        //                {
        //                    pgObj.PageNumber = page.Number;
        //                    rObj.Comments = "Text color is same.";
        //                    pglst.Add(pgObj);
        //                }
        //            }
        //            if (pglst != null && pglst.Count > 0)
        //            {
        //                rObj.CommentsPageNumLst = pglst;
        //            }
        //            if (FailedFlag != "" && PassedFlag != "" && pageNumbers == "")
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Link text color matched but border also existed for the links.";
        //            }
        //            else if (FailedFlag != "" && PassedFlag == "")
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Link border color or Text color is not in page(s):" + pageNumbers.Trim().TrimEnd(',');
        //                rObj.CommentsWOPageNum = "Link border color or Text color is not in page :";
        //                rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
        //            }
        //            else if (FailedFlag == "" && PassedFlag != "")
        //            {
        //                rObj.QC_Result = "Passed";
        //                rObj.Comments = "Text color is same.";
        //            }
        //            else if (FailedFlag != "" && PassedFlag != "")
        //            {
        //                rObj.QC_Result = "Failed";

        //                rObj.Comments = "Link border color or Text color is not in page(s):" + pageNumbers.Trim().TrimEnd(',');
        //                rObj.CommentsWOPageNum = "Link border color or Text color is not in page :";
        //                rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
        //            }
        //            else if (FailedFlag == "" && PassedFlag == "")
        //            {
        //                rObj.QC_Result = "Passed";
        //                rObj.Comments = "There are no links in the document.";
        //            }
        //        }
        //        else
        //        {
        //            rObj.QC_Result = "Failed";
        //            rObj.Comments = "There are no pages in the document";
        //        }
        //        document.Dispose();
        //        rObj.CHECK_END_TIME = DateTime.Now;
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
        //        rObj.Job_Status = "Error";
        //        rObj.QC_Result = "Error";
        //        rObj.Comments = "Technical error: " + ex.Message;


        //    }
        //}


        /// <summary>
        /// Link attributor (CBER/CDER) - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void LinkAttributor(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document document)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;
            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                rObj.CHECK_START_TIME = DateTime.Now;
                string FailedFlag = string.Empty;
                string PassedFlag = string.Empty;
                string FailedFlag1 = string.Empty;
                string PassedFlag1 = string.Empty;
                //string InvisibleTextFlag = string.Empty;
                int pgnum = 0;
                string bothExists = "";
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;

                    if (chLst[i].Check_Name.ToString() == "Text Color")
                    {
                        textcolor = chLst[i].Check_Parameter.ToString();
                    }

                    if (chLst[i].Check_Name.ToString() == "Link Border Color")
                    {
                        Linebordercolor = chLst[i].Check_Parameter.ToString();
                    }

                }

                //Document document = new Document(sourcePath);
                if (document.Pages.Count != 0)
                {

                    var editor = new PdfContentEditor(document);
                    string pageNumbers = "";
                    List<PageNumberReport> pglst = new List<PageNumberReport>();

                    // Linebordercolor = chLst[1].Check_Parameter.ToString();
                    foreach (Aspose.Pdf.Page page in document.Pages)
                    {
                        PageNumberReport pgObj = new PageNumberReport();
                        FailedFlag1 = string.Empty;
                        PassedFlag1 = string.Empty;
                        pgnum = 0;
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
                            string URL2 = string.Empty;
                            IAppointment dest = a.Destination;
                            //Checking whether link has a destination/action  or not.
                            if (a.Action != null || dest != null)
                            {

                                #region action for Aspose.Pdf.Annotations.GoToAction                                
                                if(a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                {

                                    URL2 = ((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString();
                                    if (URL2 != "")
                                    {
                                        if (dest != null)
                                        {
                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                            Aspose.Pdf.Rectangle rect = a.Rect;
                                            ta.TextSearchOptions = new TextSearchOptions(rect);
                                            ta.Visit(page);
                                            if (ta.TextFragments.Count > 0)
                                            {
                                                foreach (TextFragment tf in ta.TextFragments)
                                                {
                                                    if (tf.TextState.Invisible == false)
                                                    {
                                                        Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                                        string colortext = color1.ToString();
                                                        if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
                                                        {
                                                            FailedFlag = "Failed";
                                                            FailedFlag1 = "Failed";
                                                            pgnum = page.Number;
                                                            if (pageNumbers == "")
                                                            {
                                                                pageNumbers = page.Number.ToString() + ", ";
                                                            }
                                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                        }
                                                        else
                                                        {
                                                            int border = a.Border.Width;
                                                            Aspose.Pdf.Color lncolor = a.Color;
                                                            if (border == 0)
                                                            {
                                                                Border b = new Border(a);
                                                                b.Width = 0;
                                                                b.Style = BorderStyle.Solid;
                                                                FailedFlag = "Failed";
                                                                FailedFlag1 = "Failed";
                                                                pgnum = page.Number;
                                                                bothExists = FailedFlag;
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
                                                                PassedFlag1 = "Passed";
                                                            }
                                                                
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (Linebordercolor != "")
                                                        {
                                                            Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                                            int border = a.Border.Width;
                                                            Aspose.Pdf.Color lncolor = a.Color;
                                                            if (color.ToString().ToUpper() != lncolor.ToString().ToUpper())
                                                            {
                                                                FailedFlag = "Failed";
                                                                FailedFlag1 = "Failed";
                                                                pgnum = page.Number;
                                                                if (pageNumbers == "")
                                                                {
                                                                    pageNumbers = page.Number.ToString() + ", ";
                                                                }
                                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                            }
                                                            else
                                                            {
                                                                border = a.Border.Width;
                                                                lncolor = a.Color;
                                                                if (border == 0)
                                                                {
                                                                    Border b = new Border(a);
                                                                    b.Width = 0;
                                                                    b.Style = BorderStyle.Solid;
                                                                    FailedFlag = "Failed";
                                                                    FailedFlag1 = "Failed";
                                                                }
                                                                else
                                                                {
                                                                    PassedFlag = "Passed";
                                                                    PassedFlag1 = "Passed";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (Linebordercolor != "")
                                                {
                                                    Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                                    int border = a.Border.Width;
                                                    Aspose.Pdf.Color lncolor = a.Color;
                                                    if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
                                                    {
                                                        FailedFlag = "Failed";
                                                        FailedFlag1 = "Failed";
                                                        pgnum = page.Number;
                                                        if (pageNumbers == "")
                                                        {
                                                            pageNumbers = page.Number.ToString() + ", ";
                                                        }
                                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                    }
                                                    else
                                                    {
                                                        border = a.Border.Width;
                                                        lncolor = a.Color;
                                                        if (border == 0)
                                                        {
                                                            Border b = new Border(a);
                                                            b.Width = 0;
                                                            b.Style = BorderStyle.Solid;
                                                            FailedFlag = "Failed";
                                                            FailedFlag1 = "Failed";
                                                        }
                                                        else
                                                        {
                                                            PassedFlag = "Passed";
                                                            PassedFlag1 = "Passed";
                                                        }
                                                    }
                                                }
                                            }


                                        }
                                        else
                                        {
                                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                                            Aspose.Pdf.Rectangle rect = a.Rect;
                                            ta1.TextSearchOptions = new TextSearchOptions(rect);
                                            ta1.Visit(page);
                                            if (ta1.TextFragments.Count > 0)
                                            {
                                                foreach (TextFragment tf in ta1.TextFragments)
                                                {
                                                    string txt = tf.Text;
                                                    if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                                    {
                                                        if (tf.TextState.Invisible == false)
                                                        {
                                                            Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                                            string colortext = color1.ToString();
                                                            if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
                                                            {
                                                                FailedFlag = "Failed";
                                                                FailedFlag1 = "Failed";
                                                                pgnum = page.Number;
                                                                if (pageNumbers == "")
                                                                {
                                                                    pageNumbers = page.Number.ToString() + ", ";
                                                                }
                                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                            }
                                                            else
                                                            {
                                                                int border = a.Border.Width;
                                                                Aspose.Pdf.Color lncolor = a.Color;
                                                                if (border == 0)
                                                                {
                                                                    Border b = new Border(a);
                                                                    b.Width = 0;
                                                                    b.Style = BorderStyle.Solid;
                                                                    FailedFlag = "Failed";
                                                                    FailedFlag1 = "Failed";
                                                                    pgnum = page.Number;
                                                                    bothExists = FailedFlag;
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
                                                                    PassedFlag1 = "Passed";
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (Linebordercolor != "")
                                                            {
                                                                Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                                                int border = a.Border.Width;
                                                                Aspose.Pdf.Color lncolor = a.Color;
                                                                if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
                                                                {
                                                                    FailedFlag = "Failed";
                                                                    FailedFlag1 = "Failed";
                                                                    pgnum = page.Number;
                                                                    if (pageNumbers == "")
                                                                    {
                                                                        pageNumbers = page.Number.ToString() + ", ";
                                                                    }
                                                                    else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                                }
                                                                else
                                                                {
                                                                    border = a.Border.Width;
                                                                    lncolor = a.Color;
                                                                    if (border == 0)
                                                                    {
                                                                        Border b = new Border(a);
                                                                        b.Width = 0;
                                                                        b.Style = BorderStyle.Solid;
                                                                        FailedFlag = "Failed";
                                                                        FailedFlag1 = "Failed";
                                                                    }
                                                                    else
                                                                    {
                                                                        PassedFlag = "Passed";
                                                                        PassedFlag1 = "Passed";
                                                                        //InvisibleTextFlag = "Yes";
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (Linebordercolor != "")
                                                {
                                                    Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                                    int border = a.Border.Width;
                                                    Aspose.Pdf.Color lncolor = a.Color;
                                                    if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
                                                    {
                                                        FailedFlag = "Failed";
                                                        FailedFlag1 = "Failed";
                                                        pgnum = page.Number;
                                                        if (pageNumbers == "")
                                                        {
                                                            pageNumbers = page.Number.ToString() + ", ";
                                                        }
                                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                    }
                                                    else
                                                    {
                                                        border = a.Border.Width;
                                                        lncolor = a.Color;
                                                        if (border == 0)
                                                        {
                                                            Border b = new Border(a);
                                                            b.Width = 0;
                                                            b.Style = BorderStyle.Solid;
                                                            FailedFlag = "Failed";
                                                            FailedFlag1 = "Failed";
                                                        }
                                                        else
                                                        {
                                                            PassedFlag = "Passed";
                                                            PassedFlag1 = "Passed";
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                                #endregion
                                #region action is null                                
                                if (dest != null)
                                {
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Aspose.Pdf.Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(rect);
                                    ta.Visit(page);
                                    if (ta.TextFragments.Count > 0)
                                    {
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            if (tf.TextState.Invisible == false)
                                            {
                                                Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                                string colortext = color1.ToString();
                                                if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
                                                {
                                                    FailedFlag = "Failed";
                                                    FailedFlag1 = "Failed";
                                                    pgnum = page.Number;
                                                    if (pageNumbers == "")
                                                    {
                                                        pageNumbers = page.Number.ToString() + ", ";
                                                    }
                                                    else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                }
                                                else
                                                {
                                                    int border = a.Border.Width;
                                                    Aspose.Pdf.Color lncolor = a.Color;
                                                    if (border == 0)
                                                    {
                                                        Border b = new Border(a);
                                                        b.Width = 0;
                                                        b.Style = BorderStyle.Solid;
                                                        FailedFlag = "Failed";
                                                        FailedFlag1 = "Failed";
                                                        pgnum = page.Number;
                                                        bothExists = FailedFlag;
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
                                                        PassedFlag1 = "Passed";
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (Linebordercolor != "")
                                                {
                                                    Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                                    int border = a.Border.Width;
                                                    Aspose.Pdf.Color lncolor = a.Color;
                                                    if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
                                                    {
                                                        FailedFlag = "Failed";
                                                        FailedFlag1 = "Failed";
                                                        pgnum = page.Number;
                                                        if (pageNumbers == "")
                                                        {
                                                            pageNumbers = page.Number.ToString() + ", ";
                                                        }
                                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                    }
                                                    else
                                                    {
                                                        border = a.Border.Width;
                                                        lncolor = a.Color;
                                                        if (border == 0)
                                                        {
                                                            Border b = new Border(a);
                                                            b.Width = 0;
                                                            b.Style = BorderStyle.Solid;
                                                            FailedFlag = "Failed";
                                                            FailedFlag1 = "Failed";
                                                        }
                                                        else
                                                        {
                                                            PassedFlag = "Passed";
                                                            PassedFlag1 = "Passed";
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (Linebordercolor != "")
                                        {
                                            Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                            int border = a.Border.Width;
                                            Aspose.Pdf.Color lncolor = a.Color;
                                            if (color.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
                                            {
                                                FailedFlag = "Failed";
                                                FailedFlag1 = "Failed";
                                                pgnum = page.Number;
                                                if (pageNumbers == "")
                                                {
                                                    pageNumbers = page.Number.ToString() + ", ";
                                                }
                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                            }
                                            else
                                            {
                                                border = a.Border.Width;
                                                lncolor = a.Color;
                                                if (border == 0)
                                                {
                                                    Border b = new Border(a);
                                                    b.Width = 0;
                                                    b.Style = BorderStyle.Solid;
                                                    FailedFlag = "Failed";
                                                    FailedFlag1 = "Failed";
                                                }
                                                else
                                                {
                                                    PassedFlag = "Passed";
                                                    PassedFlag1 = "Passed";
                                                }
                                            }
                                        }
                                    }
                                }

                                #endregion
                            }
                        }
                        //if (FailedFlag1 != "" && PassedFlag1 != "" && pgnum ==0)
                        //{
                        //    pgObj.PageNumber = page.Number;
                        //    pgObj.Comments = "Link text color matched but border also existed for the links.";
                        //    pglst.Add(pgObj);
                        //}
                        //else if (FailedFlag1 != "" )
                        //{
                        //    pgObj.PageNumber = page.Number;
                        //    rObj.Comments = "No Link border color or Text color.";
                        //    pglst.Add(pgObj);
                        //}                        
                        //else if (FailedFlag1 == "" && PassedFlag1 != "")
                        //{
                        //    pgObj.PageNumber = page.Number;
                        //    rObj.Comments = "Text color is same.";
                        //    pglst.Add(pgObj);
                        //}
                    }
                    if (pglst != null && pglst.Count > 0)
                    {
                        rObj.CommentsPageNumLst = pglst;
                    }
                    if (FailedFlag != "" && PassedFlag != "" && pageNumbers=="")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Link text color matched but border also exists for the links";
                    }
                    else if (FailedFlag != "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Link border color or Text color is not in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Link border color or Text color is not";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else if (FailedFlag == "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Text color or Link border color is same.";
                    }
                    else if (FailedFlag != "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Failed";

                        rObj.Comments = "Link border color or Text color is not in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Link border color or Text color is not";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else if (FailedFlag == "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "There are no links in the document.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }                
                //document.Dispose();
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
        /// Link attributor (CBER/CDER) - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void LinkAttributorFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document document)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;
            try
            {
                //Document document = new Document(sourcePath);
                if (document.Pages.Count != 0)
                {

                    var editor = new PdfContentEditor(document);
                    string pageNumbers = "";
                    string borderPageNumbers = string.Empty;
                    // to get sub check list
                    chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                    for (int i = 0; i < chLst.Count; i++)
                    {
                        chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[i].JID = rObj.JID;
                        chLst[i].Job_ID = rObj.Job_ID;
                        chLst[i].Folder_Name = rObj.Folder_Name;
                        chLst[i].File_Name = rObj.File_Name;
                        chLst[i].Created_ID = rObj.Created_ID;

                        if (chLst[i].Check_Name.ToString() == "Text Color")
                        {
                            textcolor = chLst[i].Check_Parameter.ToString();
                        }

                        if (chLst[i].Check_Name.ToString() == "Link Border Color")
                        {
                            Linebordercolor = chLst[i].Check_Parameter.ToString();
                        }

                    }

                    if (chLst[0].Check_Type == 1)
                    {
                        rObj.CHECK_START_TIME = DateTime.Now;
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
                                bool isInvisibleExists = false;
                                string URL = string.Empty;
                                string URL1 = string.Empty; string URL2 = string.Empty;
                                IAppointment dest = a.Destination;
                                if (a.Action != null || dest != null)
                                {
                                    if (a.Action != null)
                                    {
                                        #region action as Aspose.Pdf.Annotations.GoToAction

                                        if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                        {
                                            URL2 = ((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString();
                                            if (URL2 != "")
                                            {
                                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                Aspose.Pdf.Rectangle rect = a.Rect;

                                                ta.TextSearchOptions = new TextSearchOptions(rect);
                                                ta.Visit(page);
                                                if (textcolor != "" && ta.TextFragments.Count > 0)
                                                {
                                                    foreach (TextFragment tfTemp in ta.TextFragments)
                                                    {
                                                        if (tfTemp.TextState.Invisible == true)
                                                        {
                                                            isInvisibleExists = true;
                                                            break;
                                                        }
                                                    }
                                                    if (isInvisibleExists == false)
                                                    {
                                                        foreach (TextFragment tf in ta.TextFragments)
                                                        {
                                                            string txt = tf.Text;
                                                            if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                                            //if (tf.Text.Trim() != "")
                                                            {
                                                                if (tf.TextState.Invisible == false)
                                                                {
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
                                                        //check applied to text color so need to remove line border color if it conains                                                                     
                                                        int border = a.Border.Width;
                                                        Aspose.Pdf.Color lncolor = a.Color;
                                                        if (border != 0)
                                                        {
                                                            Border b = new Border(a);
                                                            b.Width = 0;
                                                            b.Style = BorderStyle.Solid;
                                                            if (pageNumbers == "")
                                                            {
                                                                pageNumbers = page.Number.ToString() + ", ";
                                                            }
                                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";

                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (Linebordercolor != "")
                                                        {
                                                            Aspose.Pdf.Color colorLine = GetColor(Linebordercolor);
                                                            int border = a.Border.Width;
                                                            Aspose.Pdf.Color lncolor = a.Color;
                                                            if (colorLine.ToString().ToUpper() != lncolor.ToString().ToUpper()||border==0)
                                                            {
                                                                Border b = new Border(a);
                                                                b.Width = 1;
                                                                a.Color = colorLine;
                                                                b.Style = BorderStyle.Solid;                                                                                                                                
                                                                TextFixedFlag = "Fixed";
                                                            }
                                                            if (pageNumbers == "")
                                                            {
                                                                pageNumbers = page.Number.ToString() + ", ";
                                                            }
                                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    if (Linebordercolor != "")
                                                    {
                                                        Aspose.Pdf.Color colorLine = GetColor(Linebordercolor);
                                                        int border = a.Border.Width;
                                                        Aspose.Pdf.Color lncolor = a.Color;
                                                        // create TextAbsorber object to extract text
                                                        //Aspose.Pdf.Rectangle rect = a.Rect;
                                                        // create TextAbsorber object to extract text
                                                        TextAbsorber absorber = new TextAbsorber();
                                                        absorber.TextSearchOptions.LimitToPageBounds = true;
                                                        absorber.TextSearchOptions.Rectangle = rect;
                                                        // accept the absorber for first page
                                                        page.Accept(absorber);
                                                        // get the extracted text
                                                        string extractedText = absorber.Text;
                                                        if (colorLine.ToString().ToUpper() != lncolor.ToString().ToUpper() || border == 0)
                                                        {
                                                            Border b = new Border(a);
                                                            b.Width = 1;
                                                            a.Color = colorLine;
                                                            b.Style = BorderStyle.Solid;
                                                            TextFixedFlag = "Fixed";
                                                        }
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
                                    #endregion
                                    #region action as destination is not null
                                    if (dest != null)
                                    {
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Aspose.Pdf.Rectangle rect = a.Rect;

                                        ta.TextSearchOptions = new TextSearchOptions(rect);
                                        ta.Visit(page);

                                        foreach (TextFragment tfTemp in ta.TextFragments)
                                        {
                                            if (tfTemp.TextState.Invisible == true)
                                            {
                                                isInvisibleExists = true;
                                                break;
                                            }
                                        }
                                        if (isInvisibleExists == false)
                                        {
                                            foreach (TextFragment tf in ta.TextFragments)
                                            {
                                                string txt = tf.Text;
                                                if (txt.Trim() != "")
                                                {
                                                    if (textcolor != "" && tf.TextState.Invisible == false)
                                                    {
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
                                            //check applied to text color so need to remove line border color if it conains                                                                     
                                            int border = a.Border.Width;
                                            Aspose.Pdf.Color lncolor = a.Color;
                                            if (border != 0)
                                            {
                                                Border b = new Border(a);
                                                b.Width = 0;
                                                b.Style = BorderStyle.Solid;
                                                if (pageNumbers == "")
                                                {
                                                    pageNumbers = page.Number.ToString() + ", ";
                                                }
                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                            }
                                        }
                                        else
                                        {
                                            if (Linebordercolor != "")
                                            {
                                                Aspose.Pdf.Color colorLine = GetColor(Linebordercolor);
                                                int border = a.Border.Width;
                                                Aspose.Pdf.Color lncolor = a.Color;
                                                ;
                                                if (colorLine.ToString().ToUpper() != lncolor.ToString().ToUpper() || border==0)
                                                {
                                                    Border b = new Border(a);
                                                    b.Width = 1;
                                                    a.Color = colorLine;
                                                    b.Style = BorderStyle.Solid;
                                                    TextFixedFlag = "Fixed";
                                                }
                                                if (pageNumbers == "")
                                                {
                                                    pageNumbers = page.Number.ToString() + ", ";
                                                }
                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }
                            page.FreeMemory();
                        }
                        if (TextFixedFlag != "" && TextPassedFlag == "")
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = rObj.Comments + ". Fixed";
                            rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                        }
                        else if (TextFixedFlag != "" && TextPassedFlag != "")
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = rObj.Comments + ". Fixed";
                            rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                        }
                        else
                        {
                            rObj.Is_Fixed = 1;
                            rObj.Comments = rObj.Comments + ". Fixed";
                            rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                        }
                        //document.Save(sourcePath);
                        //document.Dispose();
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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
        /// No track changes - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void checkForannotationsnewbak(RegOpsQC rObj, string path)
        {
            rObj.CHECK_START_TIME = DateTime.Now;
            sourcePath = path + "//" + rObj.File_Name;
            string result = string.Empty;
            string finalResult = string.Empty;

            try
            {
                Document PdfDoc = new Document(sourcePath);
                foreach (Aspose.Pdf.Page page in PdfDoc.Pages)
                {
                    result = string.Empty;
                    foreach (Annotation annotation in page.Annotations)
                    {
                        if (annotation.AnnotationType != AnnotationType.Link)
                        {
                            if (annotation.Flags.ToString().ToLower().Contains("hidden"))
                            {
                                if (!result.Contains(page.Number.ToString() + " (Hidden),"))
                                    result = result + page.Number.ToString() + " (Hidden), ";
                            }
                            else if (!result.Contains(page.Number.ToString() + ","))
                                result = result + page.Number.ToString() + ",";
                        }
                    }
                    if (result != "")
                    {
                        finalResult = finalResult + result.Trim().TrimEnd(',') +" ," ;
                    }
                    page.FreeMemory();
                }
                if (finalResult != "")
                    finalResult = finalResult.Trim().TrimEnd(';');
                if (PdfDoc.Pages.Count == 0)
                {
                    rObj.Comments = "There are no pages in the document";
                    rObj.QC_Result = "Failed";
                }
                else if (finalResult != "")
                {
                    rObj.Comments = "Track changes(annotations) existed in the following pages " + finalResult;
                    rObj.QC_Result = "Failed";
                }
                else if (finalResult == "")
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


        /// <summary>
        /// No track changes - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void checkForannotations(RegOpsQC rObj, string path,Document PdfDoc)
        {
            rObj.CHECK_START_TIME = DateTime.Now;
            sourcePath = path + "//" + rObj.File_Name;
            string result = string.Empty;
            string finalResult = string.Empty;
            int flag = 0;
            try
            {
                //Document PdfDoc = new Document(sourcePath);
                List<PageNumberReport> pglst = new List<PageNumberReport>();
                foreach (Aspose.Pdf.Page page in PdfDoc.Pages)
                {
                    flag = 0;
                    PageNumberReport pgObj = new PageNumberReport();
                    result = string.Empty;
                    foreach (Annotation annotation in page.Annotations)
                    {
                        if (annotation.AnnotationType != AnnotationType.Link)
                        {
                            flag = 1;
                            if (annotation.Flags.ToString().ToLower().Contains("hidden"))
                            {
                                flag = 2;
                                if (!result.Contains(page.Number.ToString() + " (Hidden),"))
                                    result = result + page.Number.ToString() + " (Hidden), ";
                            }
                            else if (!result.Contains(page.Number.ToString() + ","))
                                result = result + page.Number.ToString() + ",";
                        }
                    }
                    if (result != "")
                    {
                        finalResult = finalResult + result.Trim().TrimEnd(',') + " ,";
                    }
                    if(flag == 2)
                    {
                        pgObj.PageNumber = page.Number;
                        pgObj.Comments = "Track changes exist(Hidden)";
                        pglst.Add(pgObj);
                    }
                    else if(flag == 1)
                    {
                        pgObj.PageNumber = page.Number;
                        pgObj.Comments = "Track changes exist";
                        pglst.Add(pgObj);
                    }
                    page.FreeMemory();
                }
                if (finalResult != "")
                    finalResult = finalResult.Trim().TrimEnd(';');
                if (PdfDoc.Pages.Count == 0)
                {
                    rObj.Comments = "There are no pages in the document";
                    rObj.QC_Result = "Failed";
                }
                else if (finalResult != "")
                {
                    rObj.Comments = "Track changes existed in: " + finalResult;
                    rObj.QC_Result = "Failed";
                    rObj.CommentsPageNumLst = pglst;
                }
                else if (finalResult == "")
                {
                    //rObj.Comments = "No track changes found";
                    rObj.QC_Result = "Passed";
                }                
                //PdfDoc.Dispose();
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
        /// No track changes - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void checkForannotationsFix(RegOpsQC rObj, string path,Document PdfDoc)
        {
            bool flag = false;
            rObj.FIX_START_TIME = DateTime.Now;
            sourcePath = path + "//" + rObj.File_Name;
            try
            {
                //Document PdfDoc = new Document(sourcePath);
                if (PdfDoc.Pages.Count != 0)
                {
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
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = rObj.Comments + ". Fixed";
                        if (rObj.CommentsPageNumLst != null)
                        {
                            foreach (var pg in rObj.CommentsPageNumLst)
                            {
                                pg.Comments = pg.Comments + ". Fixed";
                            }
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }
                //PdfDoc.Save(sourcePath);
                //PdfDoc.Dispose();
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
        /// Remove redundant bookmarks - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void RemoveRedundantBookmarks(RegOpsQC rObj, string path,Document document)
        {
            try
            {
                bool isRedundantBkExisted = false;
                List<string> lstbookmarks = new List<string>();
                rObj.CHECK_START_TIME = DateTime.Now;
                sourcePath = path + "//" + rObj.File_Name;
                string Result = string.Empty;
                string FinalResult = string.Empty;
                //Document document = new Document(sourcePath);
                if (document.Pages.Count != 0)
                {
                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(sourcePath);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                    //bool flag = true;
                    if (bookmarks.Count > 0)
                    {
                        for (int i = 0; i < bookmarks.Count; i++)
                        {
                            if (!lstbookmarks.Contains(bookmarks[i].Title + "_" + bookmarks[i].PageNumber))
                            {
                                lstbookmarks.Add(bookmarks[i].Title + "_" + bookmarks[i].PageNumber);
                            }
                            else
                            {
                                isRedundantBkExisted = true;
                                Result = Result + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;

                            }
                        }
                        if (isRedundantBkExisted)
                        {
                            rObj.QC_Result = "Failed";
                            FinalResult = Result.Trim().TrimStart(',');
                            rObj.Comments = "The following are the redundant bookmarks: " + FinalResult;
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "No redundant bookmarks exist in the document.";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "No bookmarks exist in the document";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }                
                //document.Dispose();
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
        /// verify bookmarks present -only check
        public void VerifyBookmarks(RegOpsQC rObj ,Document document)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                if (document.Pages.Count >= 5)
                {
                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(document);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                    if (bookmarks.Count > 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "bookmarks exist in the document.";
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No bookmarks in document";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no more than 5 pages in document";
                }
                //document.Dispose();
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
        /// Use proper special characters in Bookmarks and TOC - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckIncorrectlyConvertedSpecialCharacters(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            string FinalResult = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool isbookmarkExisted = false;
            //bool nodesbookmarks = false;
            //bool nodestoc = false;
            string nodesbkmrksres = string.Empty;
            string nodestocres = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                //Document pdfDocument = new Document(sourcePath);
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                string Result = string.Empty;
                if (bookmarks.Count > 0)
                {
                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        //if(bookmarks[i].Level>2)
                        //{
                        //}
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

                                if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                    Result = Result + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;

                                //if (Result.Length > 3800)
                                //{
                                //    int index = Result.LastIndexOf(", ");
                                //    Result = Result.Substring(0, index).TrimEnd(',');
                                //    Result = Result + " and more...";
                                //    break;
                                //}
                            }
                        }
                        else
                        {
                            if (bookmarks[i].PageNumber == 0)
                            {
                                // nodesbookmarks = true;
                                nodesbkmrksres = nodesbkmrksres + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;
                            }
                        }

                    }
                    FinalResult = Result.Trim().TrimStart(',');
                    nodesbkmrksres = nodesbkmrksres.Trim().TrimStart(',');
                    if (FinalResult != "" && nodesbkmrksres == "")
                    {
                        isbookmarkExisted = true;
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Incorrect special characters found in bookmarks as follows: " + FinalResult;
                    }
                    else if (FinalResult == "" && nodesbkmrksres != "")
                    {
                        isbookmarkExisted = true;
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks found without destination as follows: " + nodesbkmrksres;
                    }
                    else if (FinalResult != "" && nodesbkmrksres != "")
                    {
                        isbookmarkExisted = true;
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Incorrect special characters found in bookmarks as follows: " + FinalResult;
                        rObj.Comments = rObj.Comments + " Bookmarks found without destination as follows:  " + nodesbkmrksres;
                    }
                    else
                    {
                        isbookmarkExisted = true;
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No incorrect special characters found in the bookmarks";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No bookmarks existed in the document";
                }

                //Checking TOC for special charecters

                string ResultForTOC = string.Empty;
                bool isTOCLinksexisted = false;
                int TOCLinks = 0;
                bool isTOCExisted = false;
                int tocStartPageNo = 0;
                TextFragmentAbsorber textFragmentAbsorber1 = new TextFragmentAbsorber();

                Regex regex = new Regex(@"(Table of Content|Contents)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                TextFragmentAbsorber textbsorber = new TextFragmentAbsorber(regex);
                Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                textbsorber.TextSearchOptions = textSearchOptions;

                for (int pgNo = 1; pgNo <= pdfDocument.Pages.Count; pgNo++)
                {
                    pdfDocument.Pages[pgNo].Accept(textbsorber);
                    TextFragmentCollection txtFrgCollection = textbsorber.TextFragments;
                    if (txtFrgCollection.Count > 0)
                    {
                        regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        textbsorber = new TextFragmentAbsorber(regex);
                        textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                        textbsorber.TextSearchOptions = textSearchOptions;
                        pdfDocument.Pages[pgNo].Accept(textbsorber);
                        txtFrgCollection = textbsorber.TextFragments;
                        if (txtFrgCollection.Count > 0)
                        {
                            isTOCExisted = true;
                            tocStartPageNo = pgNo;
                            isTOCLinksexisted = true;
                            break;
                        }
                        else
                        {
                            for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                            {
                                Page curntPage = pdfDocument.Pages[p];
                                AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(curntPage, Aspose.Pdf.Rectangle.Trivial));
                                curntPage.Accept(selector);
                                // Create list holding all the links
                                IList<Annotation> list = selector.Selected;
                                // Iterate through invidiaul item inside list
                                foreach (LinkAnnotation a in list)
                                {
                                    string title = string.Empty;
                                    if (a.Action is GoToAction)
                                    {
                                        using (MemoryStream textStream = new MemoryStream())
                                        {
                                            // Create text device
                                            TextDevice textDevice = new TextDevice();
                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                            GoToAction linkInfo = (GoToAction)a.Action;

                                            textDevice.ExtractionOptions = textExtOptions;

                                            TextAbsorber absorber = new TextAbsorber();
                                            absorber.TextSearchOptions.Rectangle = new Aspose.Pdf.Rectangle(a.Rect.LLX, a.Rect.LLY, a.Rect.URX, a.Rect.URY);

                                            //Accept the absorber for first page
                                            pdfDocument.Pages[p].Accept(absorber);

                                            title = absorber.Text;
                                            Regex rx = new Regex(@"[.]{2,}\s?\d{1,}");
                                            if (rx.IsMatch(title))
                                            {
                                                isTOCExisted = true;
                                                tocStartPageNo = pgNo;
                                                isTOCLinksexisted = true;
                                                break;
                                            }
                                            else if (Regex.IsMatch(title, @"[.]{2,}\s?\d"))
                                            {
                                                isTOCExisted = true;
                                                tocStartPageNo = pgNo;
                                                isTOCLinksexisted = true;
                                                break;
                                            }
                                        }
                                    }
                                    else if (a.Destination != null)
                                    {
                                        using (MemoryStream textStream = new MemoryStream())
                                        {
                                            // Create text device
                                            TextDevice textDevice = new TextDevice();
                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);

                                            textDevice.ExtractionOptions = textExtOptions;

                                            TextAbsorber absorber = new TextAbsorber();
                                            absorber.TextSearchOptions.Rectangle = new Aspose.Pdf.Rectangle(a.Rect.LLX, a.Rect.LLY, a.Rect.URX, a.Rect.URY);

                                            //Accept the absorber for first page
                                            pdfDocument.Pages[p].Accept(absorber);

                                            title = absorber.Text;
                                            Regex rx = new Regex(@"[.]{2,}\s?\d{1,}");
                                            if (rx.IsMatch(title))
                                            {
                                                isTOCExisted = true;
                                                tocStartPageNo = pgNo;
                                                isTOCLinksexisted = true;
                                                break;
                                            }
                                            else if (Regex.IsMatch(title, @"[.]{2,}\s?\d"))
                                            {
                                                isTOCExisted = true;
                                                tocStartPageNo = pgNo;
                                                isTOCLinksexisted = true;
                                                break;
                                            }
                                        }
                                    }
                                }
                                if (isTOCExisted)
                                    break;
                            }
                        }
                    }
                    if (isTOCExisted)
                        break;
                }
                if (isTOCExisted && tocStartPageNo != 0)
                {
                    Document pdfDocument_Temp = new Document(sourcePath);
                    for (int p = tocStartPageNo; p <= pdfDocument.Pages.Count; p++)
                    {
                        bool noLinksExisted = true, NotocFormate = true;
                        if (isTOCExisted)
                        {
                            Page curntPage = pdfDocument.Pages[p];
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(curntPage, Aspose.Pdf.Rectangle.Trivial));
                            //using (curntPage)
                            //{
                            curntPage.Accept(selector);
                            // Create list holding all the links
                            IList<Annotation> list = selector.Selected;
                            // Iterate through invidiaul item inside list
                            foreach (LinkAnnotation a in list)
                            {
                                string title = string.Empty;
                                if (a.Action is GoToAction)
                                {
                                    TOCLinks = TOCLinks + 1;
                                    using (MemoryStream textStream = new MemoryStream())
                                    {
                                        // Create text device
                                        TextDevice textDevice = new TextDevice();
                                        // Set text extraction options - set text extraction mode (Raw or Pure)
                                        Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                        GoToAction linkInfo = (GoToAction)a.Action;

                                        textDevice.ExtractionOptions = textExtOptions;

                                        XYZExplicitDestination xyzDest = null;
                                        FitExplicitDestination fitXyz = null;
                                        FitHExplicitDestination fitHexp = null;
                                        FitBHExplicitDestination fitBHexp = null;
                                        FitVExplicitDestination fitVexp = null;
                                        FitRExplicitDestination fitRexp = null;
                                        NamedDestination namedDes = null;
                                        try
                                        {
                                            //linkInfo.Destination
                                           // xyzDest = linkInfo.Destination;
                                            xyzDest = (XYZExplicitDestination)linkInfo.Destination;
                                            noLinksExisted = false;
                                            isTOCLinksexisted = true;
                                        }
                                        catch (Exception ee)
                                        {
                                            if (ee.Message == "Unable to cast object of type 'Aspose.Pdf.Annotations.FitExplicitDestination' to type 'Aspose.Pdf.Annotations.XYZExplicitDestination'.")
                                            {
                                                fitXyz = (FitExplicitDestination)linkInfo.Destination;
                                            }
                                            else if (ee.Message == "Unable to cast object of type 'Aspose.Pdf.Annotations.FitHExplicitDestination' to type 'Aspose.Pdf.Annotations.XYZExplicitDestination'.")
                                            {
                                                fitHexp = (FitHExplicitDestination)linkInfo.Destination;
                                            }
                                            else if (ee.Message == "Unable to cast object of type 'Aspose.Pdf.Annotations.FitBHExplicitDestination' to type 'Aspose.Pdf.Annotations.XYZExplicitDestination'.")
                                            {
                                                fitBHexp = (FitBHExplicitDestination)linkInfo.Destination;
                                            }
                                            else if (ee.Message == "Unable to cast object of type 'Aspose.Pdf.Annotations.FitVExplicitDestination' to type 'Aspose.Pdf.Annotations.XYZExplicitDestination'.")
                                            {
                                                fitVexp = (FitVExplicitDestination)linkInfo.Destination;
                                            }
                                            else if (ee.Message == "Unable to cast object of type 'Aspose.Pdf.Annotations.FitRExplicitDestination' to type 'Aspose.Pdf.Annotations.XYZExplicitDestination'.")
                                            {
                                                fitRexp = (FitRExplicitDestination)linkInfo.Destination;
                                            }
                                            else if (ee.Message == "Unable to cast object of type 'Aspose.Pdf.Annotations.NamedDestination' to type 'Aspose.Pdf.Annotations.XYZExplicitDestination'.")
                                            {
                                                namedDes = (NamedDestination)linkInfo.Destination;
                                            }
                                            noLinksExisted = false;
                                            isTOCLinksexisted = true;
                                        }

                                        TextAbsorber absorber = new TextAbsorber();
                                        absorber.TextSearchOptions.Rectangle = new Aspose.Pdf.Rectangle(a.Rect.LLX, a.Rect.LLY, a.Rect.URX, a.Rect.URY);

                                        //Accept the absorber for first page

                                        pdfDocument.Pages[p].Accept(absorber);

                                        title = absorber.Text;
                                        Regex rx = new Regex(@"[.]{2,}\s?\d{1,}");
                                        if (rx.IsMatch(title))
                                        {
                                            NotocFormate = false;
                                            Match m = rx.Match(title);
                                            title = title.Replace(m.Value, "").Trim();
                                        }
                                        else if (Regex.IsMatch(title, @"[.]{2,}\s?\d"))
                                        {
                                            NotocFormate = false;
                                            title = Regex.Replace(title, @"[.]{2,}\s?\d", "");
                                        }
                                        else if (Regex.IsMatch(title, @"[.]{2,}"))
                                        {
                                            NotocFormate = false;
                                            title = Regex.Replace(title, @"[.]{2,}", "");
                                        }
                                        title = title.Replace("\r\n", "");
                                        if (title.Trim() != "")
                                        {
                                            int pageNo = 0;
                                            if (xyzDest != null)
                                            {
                                                if (xyzDest.PageNumber != 0)
                                                {
                                                    textDevice.Process(pdfDocument_Temp.Pages[xyzDest.PageNumber], textStream);
                                                    pageNo = xyzDest.PageNumber;
                                                }
                                            }
                                            else if (fitXyz != null)
                                            {
                                                if (fitXyz.PageNumber != 0)
                                                {
                                                    textDevice.Process(pdfDocument_Temp.Pages[fitXyz.PageNumber], textStream);
                                                    pageNo = fitXyz.PageNumber;
                                                }
                                            }
                                            else if (fitHexp != null)
                                            {
                                                if (fitHexp.PageNumber != 0)
                                                {
                                                    textDevice.Process(pdfDocument_Temp.Pages[fitHexp.PageNumber], textStream);
                                                    pageNo = fitHexp.PageNumber;
                                                }
                                            }
                                            else if (fitBHexp != null)
                                            {
                                                if (fitBHexp.PageNumber != 0)
                                                {
                                                    textDevice.Process(pdfDocument_Temp.Pages[fitBHexp.PageNumber], textStream);
                                                    pageNo = fitBHexp.PageNumber;
                                                }
                                            }
                                            else if (fitVexp != null)
                                            {
                                                if (fitVexp.PageNumber != 0)
                                                {
                                                    textDevice.Process(pdfDocument_Temp.Pages[fitVexp.PageNumber], textStream);
                                                    pageNo = fitVexp.PageNumber;
                                                }
                                            }
                                            else if (fitRexp != null)
                                            {
                                                if (fitRexp.PageNumber != 0)
                                                {
                                                    textDevice.Process(pdfDocument_Temp.Pages[fitRexp.PageNumber], textStream);
                                                    pageNo = fitRexp.PageNumber;
                                                }
                                            }
                                            else if (namedDes != null)
                                            {
                                                pageNo = pdfDocument_Temp.Destinations.GetPageNumber(namedDes.Name, false);
                                                if (pageNo != 0)
                                                {
                                                    textDevice.Process(pdfDocument_Temp.Pages[pageNo], textStream);
                                                }
                                            }
                                            // Get text from memory stream
                                            if (pageNo != 0)
                                            {
                                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                                string fixedStringOne = Regex.Replace(extractedText, @"\s+", String.Empty);
                                                string fixedStringTwo = Regex.Replace(title, @"\s+", String.Empty);
                                                // Close memory stream
                                                textStream.Close();
                                                if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                {
                                                    if (!ResultForTOC.Contains(title.Trim()))
                                                        ResultForTOC = ResultForTOC + ", " + title.Trim();
                                                }
                                            }
                                            else
                                            {
                                                //nodestoc = true;
                                                if (!nodestocres.Contains(title.Trim()))
                                                    nodestocres = nodestocres + ", " + title.Trim();
                                            }
                                        }
                                    }
                                }
                            }
                            curntPage.FreeMemory();
                            //}
                        }
                        if (noLinksExisted || NotocFormate)
                            break;
                    }
                    pdfDocument_Temp.Dispose();
                }
                if (rObj.QC_Result == "Passed" && (ResultForTOC != "" && isTOCExisted && nodestocres == ""))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Incorrect special characters found in the TOC as follows: " + ResultForTOC.Trim(',');
                }
                else if (rObj.QC_Result == "Passed" && (ResultForTOC == "" && isTOCExisted && nodestocres != ""))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No destination found in the TOC as follows: " + nodestocres.Trim(',');
                }
                else if (rObj.QC_Result == "Passed" && (ResultForTOC != "" && isTOCExisted && nodestocres != ""))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Incorrect special characters found in the TOC as follows: " + ResultForTOC.Trim(',') + " and No destination found in the TOC as follows: " + nodestocres.Trim(',');
                }
                else if (rObj.QC_Result == "Failed" && (ResultForTOC != "" && isTOCExisted && nodestocres == ""))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments + "\r\nIncorrect special characters found in the TOC as follows: " + ResultForTOC.Trim(',');
                }
                else if (rObj.QC_Result == "Failed" && (ResultForTOC == "" && isTOCExisted && nodestocres != ""))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments + "\r\nNo destination found in the TOC as follows: " + nodestocres.Trim(',');
                }
                else if (rObj.QC_Result == "Failed" && (ResultForTOC != "" && isTOCExisted && nodestocres != ""))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments + "\r\nIncorrect special characters found in the TOC as follows: " + ResultForTOC.Trim(',') + " and No destination found in the TOC as follows: " + nodestocres.Trim(',');
                }
                else if (rObj.QC_Result == "Passed" && (ResultForTOC == "" && !isTOCLinksexisted && nodestocres == ""))
                {
                       rObj.Comments = rObj.Comments + " and TOC not existed in the document";
                }
                else if (rObj.QC_Result == "Passed" && (isTOCExisted && isTOCLinksexisted && ResultForTOC == "" && nodestocres == ""))
                {
                    //rObj.Comments = rObj.Comments + " and no special chararecters found in the TOC";
                }
                else if (isbookmarkExisted == false && (!isTOCLinksexisted))
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There are no bookmarks and TOC existed in the document.";
                }
                //End of Checking special charecters in TOC
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
        /// Correct the Bookmark Levels - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CorrectTheBookmarkLevels(RegOpsQC rObj, string path,Document doc)
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
                bool isTOCExisted = false;
                bool isLOTExisted = false;
                bool isLOFExisted = false;
                // Open PDF file
                bookmarkEditor.BindPdf(doc);
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
                        title = Regex.Replace(bookmarksTemp[i].Title, @"\s+", " ");

                        if (title.ToUpper() == "TABLE OF CONTENTS" && bookmarksTemp[i].Level == 1)
                        {
                            originalOrder.Add(1);
                            isTOCExisted = true;
                            levelNo = levelNo + 1;
                        }
                        else if (title.ToUpper() == "LIST OF TABLES" && bookmarksTemp[i].Level == 1)
                        {
                            originalOrder.Add(2);
                            isLOTExisted = true;
                            levelNo = levelNo + 1;
                        }
                        else if (title.ToUpper() == "LIST OF FIGURES" && bookmarksTemp[i].Level == 1)
                        {
                            originalOrder.Add(3);
                            isLOFExisted = true;
                            levelNo = levelNo + 1;
                        }
                        else
                        {
                            originalOrder.Add(4);
                        }
                        if (levelNo == 3)
                            break;
                    }
                    int cunt = originalOrder.Count;
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
                            //if (isTOCExisted==false && isLOTExisted==false && isLOFExisted==false)
                            //{
                            //    //rObj.Comments = "TOC, LOT and LOF bookmarks not found";
                            //}
                            //else
                                //rObj.Comments = "Bookmarks in the document are in the correct structure";
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Bookmarks in document are not in correct structure";
                        }
                    }
                    else if (originalOrder.Count == 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks in document are not in correct structure";
                    }
                }
                else if (bookmarksTemp.Count > 4 || (bookmarks.Count > 0 && bookmarksTemp.Count == 0))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Bookmarks in document are not in correct structure";
                }
                else if (bookmarks.Count == 0)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No bookmarks existed in the document.";
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
        /// Check Links color - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckLinksColor(RegOpsQC rObj, string path,Document document)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;
            try
            {
                //Document document = new Document(sourcePath);
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
                    rObj.Comments = "Failed in: " + pageNumbers.Trim().TrimEnd(',');
                    rObj.CommentsWOPageNum = "External links color are not as per given color";
                    rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList(); 
                }
                if (FailedFlag == "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Passed";
                }
                if (FailedFlag != "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Failed in: " + pageNumbers.Trim().TrimEnd(',');
                    rObj.CommentsWOPageNum = "External links color are not as per given color";
                    rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                }
                if (FailedFlag == "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no external links in the document.";
                }
                //editor.Dispose();                                
                //document.Dispose();
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

        public void CheckInternalLinksColor(RegOpsQC rObj, string path,Document document)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;
            try
            {
                //Document document = new Document(sourcePath);
                if (document.Pages.Count != 0)
                {
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
                            try
                            {
                                //URL1 = ((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString();
                                //  if (URL1 != "")
                                // if ((a.Destination).ToString() != "")
                                if (((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString() != "")
                                {
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Aspose.Pdf.Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(rect);
                                    ta.Visit(page);
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        textcolor = rObj.Check_Parameter.ToUpper();
                                        if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                        {
                                            Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                            string colortext = color1.ToString().ToUpper();
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
                            }
                            catch
                            {
                            }
                            try
                            {
                                if ((a.Destination).ToString() != "")
                                {
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Aspose.Pdf.Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(rect);
                                    ta.Visit(page);
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        textcolor = rObj.Check_Parameter.ToUpper();
                                        if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                        {
                                            Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                            string colortext = color1.ToString().ToUpper();
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

                            }
                            catch
                            {

                            }

                        }
                        page.FreeMemory();
                    }

                    //Releasing memory for the object.
                    editor.Dispose();

                    if (FailedFlag != "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Failed in following page numbers :" + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Links color are not as per given color";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    if (FailedFlag == "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "The document has same link colors.";
                    }
                    if (FailedFlag != "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Failed in following page numbers :" + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Links color are not as per given color";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    if (FailedFlag == "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "There is no internal links in the document.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }                                
                //document.Dispose();
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
        /// Check internal links color fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckInternalLinksColorFix(RegOpsQC rObj, string path,Document document)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;
            try
            {
                //Document document = new Document(sourcePath);
                if (document.Pages.Count != 0)
                {
                    string pageNumbers = "";
                    string FixedFlag = string.Empty;
                    string PassedFlag = string.Empty;
                    Page page = null;
                    for (int i = 1; i <= document.Pages.Count; i++)
                    {
                        page = document.Pages[i];

                        // Get the link annotations from particular page
                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                        page.Accept(selector);
                        // Create list holding all the links
                        IList<Annotation> list = selector.Selected;
                        // Iterate through invidiaul item inside list                   
                        foreach (LinkAnnotation a in list)
                        {
                            //if ((a.Destination).ToString() != "")
                            if (((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString() != "")
                            {
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Aspose.Pdf.Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(rect);
                                ta.Visit(page);

                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                    {
                                        textcolor = rObj.Check_Parameter;
                                        Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                        string colortext = color1.ToString();
                                        if (textcolor != colortext)
                                        {
                                            Aspose.Pdf.Color color = GetColor(textcolor);
                                            tf.TextState.ForegroundColor = color;
                                            //rObj.QC_Result = "Fixed";
                                            rObj.Is_Fixed = 1;
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
                            if ((a.Destination).ToString() != "")
                            {
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Aspose.Pdf.Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(rect);
                                ta.Visit(page);
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                    {
                                        textcolor = rObj.Check_Parameter;
                                        Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                        string colortext = color1.ToString();
                                        if (textcolor != colortext)
                                        {
                                            Aspose.Pdf.Color color = GetColor(textcolor);
                                            tf.TextState.ForegroundColor = color;
                                            //rObj.QC_Result = "Fixed";
                                            rObj.Is_Fixed = 1;
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
                        }
                        page.FreeMemory();
                    }
                    if (FixedFlag != "")
                    {
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = "Fixed in following page numbers :" + pageNumbers.Trim().TrimEnd(',');
                    }
                }
                else
                {
                    rObj.Comments = "There are no pages in the document";
                    rObj.QC_Result = "Failed";
                }
                //document.Save(sourcePath);
                //document.Dispose();
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
        /// Check Links color - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckLinksColorFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;
            try
            {
                Document document = new Document(sourcePath);
                string FixedFlag = string.Empty;
                string PassedFlag = string.Empty;
                Page page = null;
                //foreach (Aspose.Pdf.Page page in document.Pages)
                for (int i = 1; i <= document.Pages.Count; i++)
                {
                    page = document.Pages[i];
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
                    page.FreeMemory();
                }
                if (FixedFlag != "" && PassedFlag == "")
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". These are fixed.";
                }
                document.Save(sourcePath);
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
        /// Create TOC must for above 5 page and more - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        //public void CreateTOCFromBookmarks(RegOpsQC rObj, string path)
        //{
        //    bool isTOCExisted = false;
        //    sourcePath = path + "//" + rObj.File_Name;
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    try
        //    {
        //        Rectangle r_Size = null;
        //        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        //        List<string> bookmarkkName = new List<string>();
        //        //Open PDF file
        //        bookmarkEditor.BindPdf(sourcePath);
        //        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
        //        bool isTOCLinksexisted = false;
        //        int TOCLinks = 0;
        //        Document myDocument = new Document(sourcePath);
        //        if (myDocument.Pages.Count >= 5)
        //        {
        //            if (bookmarks.Count > 0)
        //            {
        //                //Checking whether TOC existed in the current document
        //                Document currentDocToCheckTOC = new Document(sourcePath);
        //                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"(Table of Content)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //                TextFragmentAbsorber textbsorber = new TextFragmentAbsorber(regex);
        //                Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
        //                textbsorber.TextSearchOptions = textSearchOptions;
        //                for (int i = 1; i <= currentDocToCheckTOC.Pages.Count; i++)
        //                {
        //                    using (MemoryStream textStream = new MemoryStream())
        //                    {
        //                        // Create text device
        //                        TextDevice textDevice = new TextDevice();
        //                        // Set text extraction options - set text extraction mode (Raw or Pure)
        //                        Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
        //                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
        //                        textDevice.ExtractionOptions = textExtOptions;
        //                        textDevice.Process(currentDocToCheckTOC.Pages[i], textStream);
        //                        // Close memory stream
        //                        textStream.Close();
        //                        // Get text from memory stream
        //                        string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
        //                        if (regex.IsMatch(extractedText))
        //                        {
        //                            regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //                            if (regex.IsMatch(extractedText))
        //                            {
        //                                Match m = regex.Match(extractedText);
        //                                isTOCExisted = true;
        //                                isTOCLinksexisted = true;
        //                                break;
        //                            }
        //                            else if (extractedText.ToUpper().Contains("TABLE OF CONTENT"))
        //                            {
        //                                isTOCExisted = true;
        //                                isTOCLinksexisted = true;
        //                                break;
        //                                #region Pfizer format checking commented
        //                                //AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(currentDocToCheckTOC.Pages[i], Aspose.Pdf.Rectangle.Trivial));

        //                                //currentDocToCheckTOC.Pages[i].Accept(selector);
        //                                //// Create list holding all the links
        //                                //IList<Annotation> list = selector.Selected;
        //                                //// Iterate through invidiaul item inside list                            
        //                                //foreach (LinkAnnotation a in list)
        //                                //{
        //                                //    string title = string.Empty;
        //                                //    if (a.Action is GoToAction)
        //                                //    {
        //                                //        TOCLinks = TOCLinks + 1;

        //                                //        TextFragmentAbsorber titleFragmt = null;

        //                                //        GoToAction linkInfo = (GoToAction)a.Action;

        //                                //        TextAbsorber absorber = new TextAbsorber();
        //                                //        absorber.TextSearchOptions.Rectangle = new Aspose.Pdf.Rectangle(a.Rect.LLX, a.Rect.LLY, a.Rect.URX, a.Rect.URY);

        //                                //        //Accept the absorber for first page
        //                                //        currentDocToCheckTOC.Pages[i].Accept(absorber);

        //                                //        title = absorber.Text;
        //                                //        Regex rx = new Regex(@".*\s?[.]{2,}\s?\d{1,}");
        //                                //        if (rx.IsMatch(title))
        //                                //        {
        //                                //            isTOCExisted = true;
        //                                //            isTOCLinksexisted = true;
        //                                //            break;
        //                                //        }
        //                                //        else if (Regex.IsMatch(title, @"[.]{2,}\s?\d"))
        //                                //        {
        //                                //            isTOCExisted = true;
        //                                //            isTOCLinksexisted = true;
        //                                //            break;
        //                                //        }
        //                                //    }
        //                                //    else if (a.Destination != null && ((Aspose.Pdf.Annotations.ExplicitDestination)a.Destination).PageNumber != 0)
        //                                //    {
        //                                //        TOCLinks = TOCLinks + 1;

        //                                //        TextAbsorber absorber = new TextAbsorber();
        //                                //        absorber.TextSearchOptions.Rectangle = new Aspose.Pdf.Rectangle(a.Rect.LLX, a.Rect.LLY, a.Rect.URX, a.Rect.URY);

        //                                //        //Accept the absorber for first page

        //                                //        currentDocToCheckTOC.Pages[i].Accept(absorber);

        //                                //        title = absorber.Text;
        //                                //        Regex rx = new Regex(@".*\s?[.]{2,}\s?\d{1,}");
        //                                //        if (rx.IsMatch(title))
        //                                //        {
        //                                //            isTOCExisted = true;
        //                                //            isTOCLinksexisted = true;
        //                                //            break;
        //                                //        }
        //                                //        else if (Regex.IsMatch(title, @"[.]{2,}\s?\d"))
        //                                //        {
        //                                //            isTOCExisted = true;
        //                                //            isTOCLinksexisted = true;
        //                                //            break;
        //                                //        }
        //                                //    }
        //                                //}
        //                                #endregion
        //                            }
        //                        }
        //                    }
        //                    if (i == 10)
        //                        break;
        //                    else if (isTOCExisted)
        //                        break;
        //                }
        //            }
        //            else if (bookmarks.Count > 0 && isTOCExisted)
        //            {
        //                rObj.Comments = "TOC existed in the current document";
        //                rObj.QC_Result = "Passed";
        //            }
        //            if (isTOCLinksexisted && bookmarks.Count > 0)
        //            {
        //                rObj.Comments = "TOC existed in the current document";
        //                rObj.QC_Result = "Passed";
        //            }
        //            else if (!isTOCExisted && !isTOCLinksexisted && bookmarks.Count > 0)
        //            {
        //                rObj.Comments = "TOC not existed in the document";
        //                rObj.QC_Result = "Failed";
        //            }
        //            else if (bookmarks.Count == 0)
        //            {
        //                rObj.Comments = "No bookmarks existed in the document";
        //                rObj.QC_Result = "Passed";
        //            }
        //        }
        //        else if (myDocument.Pages.Count < 5)
        //        {
        //            rObj.Comments = "Document has less than 5 pages.";
        //            rObj.QC_Result = "Passed";
        //        }
        //        rObj.CHECK_END_TIME = DateTime.Now;
        //    }
        //    catch (Exception ee)
        //    {
        //        rObj.Job_Status = "Error";
        //        rObj.QC_Result = "Error";
        //        rObj.Comments = "Technical error: " + ee.Message;
        //        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);

        //    }
        //}

        /// <summary>
        /// Create TOC must for above 5 page and more - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        //public void CreateTOCFromBookmarksFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst)
        //{
        //    sourcePath = path + "//" + rObj.File_Name;

        //    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        //    //Getting page size of the original document.
        //    PdfPageEditor pageEditor = new PdfPageEditor();
        //    pageEditor.BindPdf(sourcePath);
        //    PageSize originalPG = pageEditor.GetPageSize(1);
        //    pageEditor.Close();

        //    List<string> bookmarkkName = new List<string>();
        //    //Open PDF file
        //    bookmarkEditor.BindPdf(sourcePath);
        //    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
        //    bookmarkEditor.Close();

        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    try
        //    {
        //        bool fixornot = false;
        //        chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
        //        for (int i = 0; i < chLst.Count; i++)
        //        {
        //            if (chLst[i].Check_Type == 1)
        //            {
        //                fixornot = true;
        //                break;
        //            }
        //        }
        //        if (fixornot && bookmarks.Count > 0)
        //        {
        //            Document myDocument = new Document(sourcePath);
        //            int initialPageCount = myDocument.Pages.Count();

        //            TocInfo tocInfo = new TocInfo();
        //            Aspose.Pdf.Page tocPage = myDocument.Pages.Insert(1);

        //            //Set Page size for TOC pages
        //            if (originalPG.Width > originalPG.Height)
        //                myDocument.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
        //            else
        //                myDocument.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);

        //            TextFragment titleFrag = new TextFragment();

        //            bool tocTitleFlag = false;
        //            bool lotFlag = false;
        //            bool lofFlag = false;
        //            bool isLOTExisted = false;
        //            bool isLOFExisted = false;

        //            //If any duplicate TOC bookmark existed removing it
        //            for (int i = 0; i < bookmarks.Count(); i++)
        //            {
        //                if (bookmarks[i].Title.ToUpper() == "TABLE OF CONTENTS" && bookmarks[i].ChildItems.Count == 0)
        //                {
        //                    bookmarks.RemoveAt(i);
        //                    break;
        //                }
        //            }

        //            for (int i = 0; i < bookmarks.Count; i++)
        //            {
        //                string title = bookmarks[i].Title;

        //                if (title.ToUpper() != rObj.File_Name.ToUpper().Replace(".PDF", "") && (title.ToUpper() == "TABLE OF CONTENTS" || title.ToUpper() != "LIST OF TABLES" || title.ToUpper() != "LIST OF FIGURES"))
        //                {
        //                    if (tocTitleFlag == false && bookmarks[i].Level == 1)
        //                    {

        //                        titleFrag = new TextFragment();
        //                        titleFrag.Text = "TABLE OF CONTENTS";
        //                        titleFrag.TextState.LineSpacing = 20;
        //                        titleFrag.TextState.FontSize = 12;
        //                        titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                        titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                        titleFrag.TextState.FontStyle = FontStyles.Bold;
        //                        //tocPage.Paragraphs.Add(titleFrag);
        //                        tocInfo = new TocInfo();
        //                        tocInfo.Title = titleFrag;
        //                        tocPage.TocInfo = tocInfo;

        //                        tocTitleFlag = true;
        //                    }

        //                    rObj.Comments = "Table of contents created as per the bookmarks";
        //                    rObj.QC_Result = "Fixed";
        //                    if (title.ToUpper() != "LIST OF TABLES" && title.ToUpper() != "LIST OF FIGURES" && bookmarks[i].Level <= 4)
        //                    {
        //                        Aspose.Pdf.Heading heading2 = new Heading(1);
        //                        heading2 = SetPropertiesForTOCItemsNew(rObj, bookmarks[i].Level, chLst);

        //                        TextSegment segment2 = new TextSegment();
        //                        heading2.TocPage = tocPage;
        //                        segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst, title);
        //                        //segment2.Text = title;
        //                        heading2.TextState.ForegroundColor = Color.Blue;
        //                        heading2.TextState.LineSpacing = 10;

        //                        // Destination page
        //                        heading2.DestinationPage = myDocument.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber)]; //myDocument.Pages[i + 2];                        
        //                        heading2.Top = myDocument.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber)].Rect.Height;
        //                        //LocalHyperlink lhl = new LocalHyperlink();
        //                        //lhl.TargetPageNumber = (bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber);
        //                        //segment2.Hyperlink = lhl;

        //                        if (bookmarks[i].Level == 2)
        //                        {
        //                            heading2.Margin.Left = 15;
        //                        }
        //                        else if (bookmarks[i].Level == 3)
        //                        {
        //                            heading2.Margin.Left = 25;
        //                        }
        //                        else if (bookmarks[i].Level == 4 || bookmarks[i].Level > 4)
        //                        {
        //                            heading2.Margin.Left = 30;
        //                        }


        //                        heading2.Segments.Add(segment2);
        //                        tocPage.Paragraphs.Add(heading2);
        //                    }

        //                }
        //                if (title.ToUpper() == "LIST OF TABLES" && tocTitleFlag == true && lotFlag == false)
        //                {
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF TABLES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;
        //                    tocPage.Paragraphs.Add(titleFrag);
        //                    lotFlag = true;
        //                    isLOTExisted = true;
        //                }
        //                else if (title.ToUpper() == "LIST OF FIGURES" && tocTitleFlag == true && lofFlag == true)
        //                {
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF FIGURES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;
        //                    tocPage.Paragraphs.Add(titleFrag);
        //                    lofFlag = true;
        //                    isLOFExisted = true;
        //                }
        //                if (title.ToUpper() == "LIST OF FIGURES")
        //                    isLOFExisted = true;
        //                else if (title.ToUpper() == "LIST OF TABLES")
        //                    isLOTExisted = true;
        //            }

        //            Guid guid = Guid.NewGuid();

        //            myDocument.Save(sourcePath1 + guid + rObj.File_Name);
        //            myDocument.Dispose();

        //            myDocument = new Document(sourcePath1 + guid + rObj.File_Name);
        //            int tocPages = myDocument.Pages.Count - initialPageCount;


        //            myDocument.Dispose();
        //            //Again taking source file as input file
        //            myDocument = new Document(sourcePath);
        //            for (int n = 1; n <= tocPages; n++)
        //            {
        //                myDocument.Pages.Add();
        //                //myDocument.Pages.Insert(n);
        //            }
        //            myDocument.Save(sourcePath);
        //            myDocument.Dispose();

        //            //Set page size for newly added pages
        //            //myDocument = new Document(sourcePath);                   

        //            //for (int n = 1; n <= tocPages; n++)
        //            //{
        //            //    if (originalPG.Width > originalPG.Height)
        //            //        myDocument.Pages[n].SetPageSize(originalPG.Height, originalPG.Width);
        //            //    else
        //            //        myDocument.Pages[n].SetPageSize(originalPG.Width, originalPG.Height);
        //            //    //myDocument.Pages.Insert(n);
        //            //}
        //            //myDocument.Save(sourcePath);
        //            //myDocument.Dispose();

        //            Document myDocumentNew = new Document(sourcePath);
        //            tocInfo = new TocInfo();
        //            tocPage = myDocumentNew.Pages.Insert(1);
        //            if (originalPG.Width > originalPG.Height)
        //                myDocumentNew.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
        //            else
        //                myDocumentNew.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);
        //            //tocPage = myDocumentNew.Pages[1];

        //            tocTitleFlag = false;
        //            lotFlag = false;
        //            lofFlag = false;
        //            int lotCount = 0;
        //            int lofCount = 0;
        //            TextFragment titleFrag1 = new TextFragment();
        //            int initialPageNo = 0;
        //            int flagCount = 0;
        //            for (int i = 0; i < bookmarks.Count; i++)
        //            {
        //                string title = bookmarks[i].Title;

        //                if (tocTitleFlag == false && bookmarks[i].Level == 1 && (title.ToUpper() != "LIST OF TABLES") && title.ToUpper() != "LIST OF FIGURES")
        //                {
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "TABLE OF CONTENTS";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;
        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;

        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    tocTitleFlag = true;
        //                }
        //                else if (title.ToUpper() == "LIST OF TABLES" && tocTitleFlag == true && lotFlag == false)
        //                {
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF TABLES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;
        //                    tocPage.Paragraphs.Add(titleFrag);
        //                    lotFlag = true;
        //                    lotCount = bookmarks[i].ChildItems.Count;
        //                    lofCount = 0;
        //                    flagCount = 0;
        //                }
        //                else if (title.ToUpper() == "LIST OF FIGURES" && ((tocTitleFlag == true && lotFlag == true && isLOTExisted) || (tocTitleFlag == true && lotFlag == false && !isLOTExisted)))
        //                {
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF FIGURES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;
        //                    tocPage.Paragraphs.Add(titleFrag);
        //                    lofCount = bookmarks[i].ChildItems.Count;
        //                    lofFlag = true;
        //                    flagCount = 0;
        //                    lotCount = 0;
        //                }


        //                if (tocTitleFlag == true && title.ToUpper() != "LIST OF TABLES" && title.ToUpper() != "LIST OF FIGURES" && bookmarks[i].Level <= 4)
        //                {
        //                    if (lotCount > 0)
        //                    {
        //                        flagCount++;
        //                    }
        //                    if (lofCount > 0)
        //                    {
        //                        flagCount++;
        //                    }
        //                    Aspose.Pdf.Heading heading2 = new Heading(1);
        //                    heading2 = SetPropertiesForTOCItemsNew(rObj, bookmarks[i].Level, chLst);
        //                    //TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst);

        //                    TextSegment segment2 = new TextSegment();
        //                    heading2.TocPage = tocPage;
        //                    segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst, title);
        //                    heading2.TextState.ForegroundColor = Color.Blue;
        //                    heading2.TextState.LineSpacing = 10;

        //                    // Destination page
        //                    heading2.DestinationPage = myDocumentNew.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber) + 1]; //myDocument.Pages[i + 2];                        
        //                    heading2.Top = myDocumentNew.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber) + 1].Rect.Height;
        //                    //LocalHyperlink lhl = new LocalHyperlink();
        //                    //lhl.TargetPageNumber = (bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber) + tocPages;
        //                    //segment2.Hyperlink = lhl;

        //                    if (bookmarks[i].Level == 2)
        //                    {
        //                        heading2.Margin.Left = 15;
        //                    }
        //                    else if (bookmarks[i].Level == 3)
        //                    {
        //                        heading2.Margin.Left = 25;
        //                    }
        //                    else if (bookmarks[i].Level == 4 || bookmarks[i].Level > 4)
        //                    {
        //                        heading2.Margin.Left = 30;
        //                    }

        //                    heading2.Segments.Add(segment2);
        //                    tocPage.Paragraphs.Add(heading2);

        //                }

        //                if (lotFlag == false && isLOTExisted == true && tocTitleFlag == true && i == bookmarks.Count() - 1)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                }
        //                else if (lofFlag == false && isLOFExisted == true && tocTitleFlag == true && i == bookmarks.Count() - 1)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                }
        //                else if (lofFlag == false && isLOFExisted == true && tocTitleFlag == true && !isLOTExisted && i == bookmarks.Count() - 1)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                }

        //                //if (lotCount > 0 && flagCount == lotCount && isLOFExisted)
        //                //    i = -1;
        //                //else
        //                if (lotCount > 0 && flagCount == lotCount && !isLOFExisted)
        //                    break;
        //                else if (lofCount > 0 && flagCount == lofCount && flagCount > 0)
        //                    break;
        //            }
        //            myDocumentNew.Save(sourcePath);
        //            myDocumentNew.Dispose();

        //            bookmarkEditor = new PdfBookmarkEditor();
        //            //Open PDF file
        //            bookmarkEditor.BindPdf(sourcePath);
        //            bookmarks = bookmarkEditor.ExtractBookmarks();

        //            for (int i = 0; i < bookmarks.Count(); i++)
        //            {
        //                if (bookmarks[i].Title.ToUpper() == "TABLE OF CONTENTS" && bookmarks[i].ChildItems.Count == 0)
        //                {
        //                    bookmarks.RemoveAt(i);
        //                    break;
        //                }
        //            }

        //            myDocumentNew = new Document(sourcePath);
        //            TextFragmentAbsorber textFragmentAbsorber = null;

        //            // Get the extracted text fragments
        //            bool setTOC = false;
        //            bool setLOT = false;
        //            bool setLOF = false;

        //            bool isTocLinkCreated = false;
        //            for (int pn = 1; pn <= myDocumentNew.Pages.Count; pn++)
        //            {
        //                textFragmentAbsorber = new TextFragmentAbsorber();
        //                // Accept the absorber for all the pages
        //                myDocumentNew.Pages[pn].Accept(textFragmentAbsorber);

        //                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
        //                for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
        //                {
        //                    TextFragment tf = textFragmentCollection[frg];
        //                    if (tf.Text.ToUpper().Contains("LIST OF TABLES"))
        //                    {
        //                        Rectangle rect = tf.Rectangle;
        //                        Bookmark bookmarkLOT = new Bookmark();
        //                        //bookmarkTOC.Title = "TABLE OF CONTENTS";
        //                        for (int bk = 0; bk < bookmarks.Count; bk++)
        //                        {
        //                            if (bookmarks[bk].Title == "LIST OF TABLES")
        //                            {
        //                                bookmarkLOT = bookmarks[bk];
        //                                //bookmarkLOT.Level = 1;
        //                                bookmarkLOT.PageNumber = tf.Page.Number;
        //                                //bookmarkLOT.Action = "GoTo";
        //                                //bookmarkLOT.PageDisplay = "XYZ";
        //                                bookmarkLOT.PageDisplay_Left = (int)rect.LLX;
        //                                bookmarkLOT.PageDisplay_Top = (int)rect.URY;
        //                                //bookmarkLOT.PageDisplay_Zoom = 0;
        //                                bookmarks[bk] = bookmarkLOT;
        //                                setLOT = true;
        //                                break;
        //                            }
        //                        }
        //                    }
        //                    if (tf.Text.ToUpper().Contains("LIST OF FIGURES"))
        //                    {
        //                        Rectangle rect = tf.Rectangle;
        //                        Bookmark bookmarkLOF = new Bookmark();
        //                        //bookmarkTOC.Title = "TABLE OF CONTENTS";
        //                        for (int bk = 0; bk < bookmarks.Count; bk++)
        //                        {
        //                            if (bookmarks[bk].Title == "LIST OF FIGURES")
        //                            {
        //                                bookmarkLOF = bookmarks[bk];
        //                                //bookmarkLOF.Level = 1;
        //                                bookmarkLOF.PageNumber = tf.Page.Number;
        //                                //bookmarkLOF.Action = "GoTo";
        //                                //bookmarkLOF.PageDisplay = "XYZ";
        //                                bookmarkLOF.PageDisplay_Left = (int)rect.LLX;
        //                                bookmarkLOF.PageDisplay_Top = (int)rect.URY;
        //                                //bookmarkLOF.PageDisplay_Zoom = 0;
        //                                bookmarks[bk] = bookmarkLOF;
        //                                setLOF = true;
        //                                break;
        //                            }
        //                        }
        //                    }
        //                    if (tf.Text.ToUpper().Contains("TABLE OF CONTENTS"))
        //                    {
        //                        Rectangle rect = tf.Rectangle;
        //                        Bookmark bookmarkTOC = new Bookmark();
        //                        bookmarkTOC.Title = "TABLE OF CONTENTS";
        //                        bookmarkTOC.Level = 1;
        //                        bookmarkTOC.PageNumber = tf.Page.Number;
        //                        bookmarkTOC.Action = "GoTo";
        //                        bookmarkTOC.PageDisplay = "XYZ";
        //                        bookmarkTOC.PageDisplay_Left = (int)rect.LLX;
        //                        bookmarkTOC.PageDisplay_Top = (int)rect.URY;
        //                        bookmarkTOC.PageDisplay_Zoom = 0;
        //                        bookmarks.Insert(0, bookmarkTOC);
        //                        setTOC = true;
        //                        //isTocLinkCreated = true;
        //                        //break;
        //                    }
        //                }
        //                //if (isTocLinkCreated)
        //                //    break;
        //                if (setTOC && setLOT && setLOF)
        //                    break;
        //            }

        //            bookmarkEditor.DeleteBookmarks();

        //            for (int bk = 0; bk < bookmarks.Count; bk++)
        //            {
        //                if (bookmarks[bk].Level == 1)
        //                    bookmarkEditor.CreateBookmarks(bookmarks[bk]);
        //            }

        //            bookmarkEditor.Save(sourcePath);
        //            myDocumentNew.Dispose();

        //            //myDocument = new Document(sourcePath);
        //            //PdfPageEditor pageEditor = new PdfPageEditor();
        //            //pageEditor.BindPdf(sourcePath);
        //            //PageSize originalPG = pageEditor.GetPageSize(tocPages + 1);
        //            //PageSize pz = null;
        //            //pageEditor.Close();
        //            //if (originalPG.Width > originalPG.Height)
        //            //{
        //            //    pz = new PageSize(originalPG.Height, originalPG.Width);
        //            //    //float width = (float)8.5;
        //            //    //pz = new PageSize(width, 11);
        //            //    originalPG = pz;
        //            //}

        //            //pageEditor = new PdfPageEditor();
        //            //pageEditor.BindPdf(sourcePath);
        //            //List<int> pgList = new List<int>();
        //            //for (int p = 1; p <= tocPages; p++)
        //            //{
        //            //    pgList.Add(p);
        //            //}
        //            //pageEditor.ProcessPages = pgList.ToArray();
        //            //pageEditor.PageSize = originalPG;
        //            //pageEditor.Save(sourcePath);
        //            //pageEditor.Close();

        //            myDocumentNew = new Document(sourcePath);
        //            int extraPagesCnt = myDocumentNew.Pages.Count;
        //            for (int n = 0; n < (extraPagesCnt - (initialPageCount + tocPages)); n++)
        //            {
        //                myDocumentNew.Pages.Delete(myDocumentNew.Pages.Count);
        //            }
        //            myDocumentNew.Save(sourcePath);
        //            myDocumentNew.Dispose();
        //            rObj.Comments = "Table of contents created as per the bookmarks";
        //            rObj.QC_Result = "Fixed";
        //        }
        //        else
        //        {
        //            rObj.Comments = "No bookmarks existed in the document";
        //            rObj.QC_Result = "Passed";
        //        }
        //        rObj.CHECK_END_TIME = DateTime.Now;

        //    }
        //    catch (Exception ee)
        //    {
        //        rObj.Job_Status = "Error";
        //        rObj.QC_Result = "Error";
        //        rObj.Comments = "Technical error: " + ee.Message;
        //        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
        //    }
        //}

        public TextSegment SetPropertiesForTOCItems(RegOpsQC rObj, int Level, List<RegOpsQC> chLst, string Title)
        {
            //the below code can be used for TextSegment,TextFragment
            TextSegment textFrg = new TextSegment();
            textFrg.Text = Title + " ";
            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                chLst = chLst.Where(x => x.Check_Name.Contains(Level.ToString())).ToList();
                //applying styles
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

                        if (chLst[fs].Check_Name == "Level" + Level + " - Font Family" && chLst[fs].Check_Type == 1)
                        {
                            textFrg.TextState.Font = FontRepository.FindFont(chLst[fs].Check_Parameter);
                        }
                        else if (chLst[fs].Check_Name == "Level" + Level + " - Font Style" && chLst[fs].Check_Type == 1)
                        {
                            if (chLst[fs].Check_Parameter == "Bold")
                                textFrg.TextState.FontStyle = FontStyles.Bold;
                            else if (chLst[fs].Check_Parameter == "Italic")
                                textFrg.TextState.FontStyle = FontStyles.Italic;
                            else if (chLst[fs].Check_Parameter == "Regular")
                                textFrg.TextState.FontStyle = FontStyles.Regular;
                        }
                        else if (chLst[fs].Check_Name == "Level" + Level + " - Font Size" && chLst[fs].Check_Type == 1)
                            textFrg.TextState.FontSize = float.Parse(chLst[fs].Check_Parameter);
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

        /// <summary>
        /// Check hyperlinks auditor - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void LinkAuditorManificationGoToViewInternal(RegOpsQC rObj, string path,Document document)
        {
            try
            {
                string res = string.Empty;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;

                //Document document = new Document(sourcePath);
                if (document.Pages.Count != 0)
                {
                    int linksflag = 0;
                    bookmarksflag = 0;
                    bool linksFailedFlag = false;
                    foreach (Aspose.Pdf.Page page in document.Pages)
                    {
                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        foreach (LinkAnnotation a in list)
                        {
                            try
                            {
                                if (a.Action != null)
                                {
                                    if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                    {
                                        XYZExplicitDestination xyz = ((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination as XYZExplicitDestination;
                                        if (xyz == null || (xyz).Zoom > 0)
                                        {
                                            linksflag = 2;
                                            linksFailedFlag = true;
                                        }
                                        else if ((xyz).Zoom == 0)
                                        {
                                            linksflag = 1;
                                        }
                                    }
                                
                                }
                                else if (a.Destination != null && (a.Destination).ToString() != "")
                                {
                                    XYZExplicitDestination xyz = a.Destination as XYZExplicitDestination;
                                    if (xyz == null || (xyz).Zoom > 0)
                                    {
                                        linksflag = 2;
                                        linksFailedFlag = true;
                                    }
                                    else if ((xyz).Zoom == 0)
                                    {
                                        linksflag = 1;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ErrorLogger.Error(ex);
                            }
                        }
                        page.FreeMemory();
                    }
                    // the below code for check all bookmarks are i inherit zoom
                    int flag = 0;
                    bool FailedFlag = false;
                    List<string> lstBookmarksToBeFixed = new List<string>();
                    if (document.Outlines.Count > 0)
                    {
                        bookmarksflag = 1;
                        foreach (OutlineItemCollection outlineItem in document.Outlines)
                        {
                            try
                            {
                                if (outlineItem.Action != null)
                                {
                                    if (((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination != null && ((Aspose.Pdf.Annotations.XYZExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination).Zoom != 0)
                                    {
                                        lstBookmarksToBeFixed.Add(outlineItem.Title);
                                        FailedFlag = true;
                                        bookmarksflag = 2;

                                    }
                                }
                                else
                                {
                                    if (outlineItem.Destination != null)
                                    {
                                        if (!outlineItem.Destination.ToString().Contains("XYZ") && ((Aspose.Pdf.Annotations.XYZExplicitDestination)outlineItem.Destination).Zoom != 0)
                                        {
                                            lstBookmarksToBeFixed.Add(outlineItem.Title);
                                            FailedFlag = true;
                                            bookmarksflag = 2;
                                        }
                                    }
                                }                                
                                if (outlineItem.Count > 0)
                                {
                                    flag = RecursiveBookmarkOutlineFix(outlineItem, lstBookmarksToBeFixed);
                                    if (flag == 2)
                                    {
                                        bookmarksflag = 2;
                                        FailedFlag = true;
                                    }

                                }

                            }
                            catch
                            {
                                try
                                {
                                    if (outlineItem.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                    {
                                        if (((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination != null && (((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination.ToString().Contains("FitBH")) || ((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination.ToString().Contains("FitH") || ((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination.ToString().Contains("Fit"))
                                        {
                                            bookmarksflag = 2;
                                            lstBookmarksToBeFixed.Add(outlineItem.Title);
                                            FailedFlag = true;
                                        }
                                        if (outlineItem.Count > 0)
                                        {
                                            flag = RecursiveBookmarkOutlineFix(outlineItem, lstBookmarksToBeFixed);
                                            if (flag == 2)
                                            {
                                                bookmarksflag = 2;
                                                FailedFlag = true;
                                            }

                                        }
                                    }
                                        
                                }
                                catch
                                {
                                    if (outlineItem.Destination != null)
                                    {
                                        if (outlineItem.Destination.ToString().Contains("FitH"))
                                        {
                                            lstBookmarksToBeFixed.Add(outlineItem.Title);
                                            FailedFlag = true;
                                            bookmarksflag = 2;
                                        }
                                    }
                                    if (outlineItem.Count > 0)
                                    {
                                        flag = RecursiveBookmarkOutlineFix(outlineItem, lstBookmarksToBeFixed);
                                        if (flag == 2)
                                        {
                                            bookmarksflag = 2;
                                            FailedFlag = true;
                                        }

                                    }
                                }
                            }
                        }
                    }

                    //PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                    //pdfEditor.BindPdf(sourcePath);
                    //Bookmarks bookmarks = pdfEditor.ExtractBookmarks();
                    //int NoZoomLevelCount = 0;
                    //if (bookmarks.Count > 0)
                    //{
                    //    bookmarksflag = 1;
                    //    for (int i = 0; i < bookmarks.Count; i++)
                    //    {          
                    //        if(bookmarks[i].PageDisplay==null)
                    //        {
                    //            NoZoomLevelCount++;
                    //        }
                    //        else if (bookmarks[i].PageDisplay != "XYZ" || (bookmarks[i].PageDisplay == "XYZ" && bookmarks[i].PageDisplay_Zoom != 0))
                    //            bookmarksflag = 2;                            
                    //    }
                    //}
                    if (linksFailedFlag && bookmarksflag == 2)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Links and bookmarks are not set to inherit zoom";
                    }
                    else if (linksFailedFlag && bookmarksflag == 1 && FailedFlag == false && document.Outlines.Count > 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Links are not set to inherit zoom and bookmarks are already set to inherit zoom";
                    }
                    else if (linksFailedFlag == false && bookmarksflag == 2 && linksflag != 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Links are already set to inherit zoom and bookmarks are not set to inherit zoom";
                    }
                    else if (linksFailedFlag == false && bookmarksflag == 1 && FailedFlag == false && document.Outlines.Count > 0 && linksflag != 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Links and bookmarks are already set to inherit zoom";
                    }
                    else if (linksflag == 1 && linksFailedFlag == false && bookmarksflag == 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Links are already set to inherit zoom and no bookmarks exist in document";
                    }
                    else if (linksflag == 0 && bookmarksflag == 1 && FailedFlag == false && document.Outlines.Count > 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Bookmarks are already set to inherit zoom and no links exist in document";
                    }
                    else if (linksflag == 0 && bookmarksflag == 2)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks are not set to inherit zoom and no links exist in document";
                    }
                    else if (linksFailedFlag && bookmarksflag == 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Links are not set to inherit zoom and no bookmarks exist in document";
                    }
                    else if (linksflag == 0 && FailedFlag && document.Outlines.Count > 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks are not in inherit zoom and No links existed in the document";
                    }
                    else if (linksFailedFlag == false && linksflag != 0 && FailedFlag && document.Outlines.Count > 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks are not set to inherit zoom and Links are already set to inherit zoom";
                    }
                    else if (linksFailedFlag && linksflag != 0 && FailedFlag && document.Outlines.Count > 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks are not set to inherit zoom and Links are not set to inherit zoom";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No links and bookmarks exist in document";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }                
                //document.Dispose();
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
        public int RecursiveBookmarkOutlineFix(OutlineItemCollection outlineItem, List<string> lstBookmarksToBeFixed)
        {
            int flag = 0;
            foreach (OutlineItemCollection childOutline in outlineItem)
            {
                try
                {
                    if (childOutline.Action != null)
                    {
                        if (((Aspose.Pdf.Annotations.GoToAction)childOutline.Action).Destination != null && ((Aspose.Pdf.Annotations.XYZExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)childOutline.Action).Destination).Zoom != 0)
                        {
                            lstBookmarksToBeFixed.Add(childOutline.Title);
                            bookmarksflag = 2;
                        }
                    }
                    else
                    {
                        if (childOutline.Destination != null)
                        {
                            if (!childOutline.Destination.ToString().Contains("XYZ"))
                            {
                                try
                                {
                                    if ((Aspose.Pdf.Annotations.GoToAction)childOutline.Action != null)
                                    {
                                        if (((Aspose.Pdf.Annotations.XYZExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)childOutline.Action).Destination).Zoom != 0)
                                        {
                                            lstBookmarksToBeFixed.Add(childOutline.Title);
                                            bookmarksflag = 2;
                                        }
                                    }
                                }
                                catch
                                {
                                    lstBookmarksToBeFixed.Add(childOutline.Title);
                                    bookmarksflag = 2;
                                }

                            }
                        }
                    }
                    if (childOutline.Count > 0)
                    {
                        flag = RecursiveBookmarkOutlineFix(childOutline, lstBookmarksToBeFixed);
                    }

                }
                catch
                {
                    if (((Aspose.Pdf.Annotations.GoToAction)childOutline.Action).Destination != null && (((Aspose.Pdf.Annotations.GoToAction)childOutline.Action).Destination.ToString().Contains("FitBH")) || ((Aspose.Pdf.Annotations.GoToAction)childOutline.Action).Destination.ToString().Contains("FitH") || ((Aspose.Pdf.Annotations.GoToAction)childOutline.Action).Destination.ToString().Contains("Fit"))
                    {
                        lstBookmarksToBeFixed.Add(childOutline.Title);
                        flag = 2;
                        bookmarksflag = 2;
                    }
                    if (childOutline.Count > 0)
                    {
                        flag = RecursiveBookmarkOutlineFix(childOutline, lstBookmarksToBeFixed);
                    }
                }
            }
            return flag;
        }

        /// <summary>
        /// Check hyperlinks auditor - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void LinkAuditorManificationGoToViewInternalFix(RegOpsQC rObj, string path,Document document)
        {
            try
            {
                if (rObj.Comments != "Bookmarks are not having zoom level property and No links existed in the document")
                {
                    string res = string.Empty;
                    sourcePath = path + "//" + rObj.File_Name;
                    rObj.FIX_START_TIME = DateTime.Now;
                    //Document document = new Document(sourcePath);
                    bool linksFailedFlag = false;
                    int NoZoomLevelCount = 0;
                    if (document.Pages.Count != 0)
                    {
                        bool IsNamedDestinationsExisted = false;
                        int linksflag = 0;
                        bookmarksflag = 0;
                        int flag = 0;
                        List<string> lstBookmarksToBeFixed = new List<string>();
                        if (document.Outlines.Count > 0)
                        {
                            bookmarksflag = 1;
                            foreach (OutlineItemCollection outlineItem in document.Outlines)
                            {
                                try
                                {
                                    if (outlineItem.Action != null)
                                    {
                                        if (((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination != null && ((Aspose.Pdf.Annotations.XYZExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination).Zoom != 0)
                                        {
                                            lstBookmarksToBeFixed.Add(outlineItem.Title);
                                        }
                                    }
                                    else
                                    {
                                        if (outlineItem.Destination != null)
                                        {
                                            if (!outlineItem.Destination.ToString().Contains("XYZ") && ((Aspose.Pdf.Annotations.XYZExplicitDestination)outlineItem.Destination).Zoom != 0)
                                            {
                                                lstBookmarksToBeFixed.Add(outlineItem.Title);
                                            }
                                        }
                                    }
                                    if (outlineItem.Count > 0)
                                    {
                                        flag = RecursiveBookmarkOutlineFix(outlineItem, lstBookmarksToBeFixed);
                                    }

                                }
                                catch
                                {
                                    if (outlineItem.Action != null)
                                    {
                                        if (((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination != null && (((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination.ToString().Contains("FitBH")) || ((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination.ToString().Contains("FitH") || ((Aspose.Pdf.Annotations.GoToAction)outlineItem.Action).Destination.ToString().Contains("Fit"))
                                        {
                                            lstBookmarksToBeFixed.Add(outlineItem.Title);
                                        }
                                        if (outlineItem.Count > 0)
                                        {
                                            flag = RecursiveBookmarkOutlineFix(outlineItem, lstBookmarksToBeFixed);
                                        }
                                    }
                                    else
                                    {
                                        if ((outlineItem.Destination.ToString().Contains("FitBH")) || (outlineItem.Destination.ToString().Contains("FitH")) || (outlineItem.Destination.ToString().Contains("Fit")))
                                        {
                                            lstBookmarksToBeFixed.Add(outlineItem.Title);
                                        }
                                        if (outlineItem.Count > 0)
                                        {
                                            flag = RecursiveBookmarkOutlineFix(outlineItem, lstBookmarksToBeFixed);
                                        }
                                    }

                                }
                            }
                        }
                        if (lstBookmarksToBeFixed.Count > 0)
                        {
                            PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                            pdfEditor.BindPdf(document);
                            Bookmarks bookmarks = pdfEditor.ExtractBookmarks();

                            if (bookmarks.Count > 0)
                            {
                                bookmarksflag = 1;
                                for (int i = 0; i < bookmarks.Count; i++)
                                {
                                    if (bookmarks[i].PageDisplay == null)
                                    {
                                        NoZoomLevelCount++;
                                    }
                                    else if (lstBookmarksToBeFixed.Contains(bookmarks[i].Title))
                                    {
                                        bookmarks[i].PageDisplay_Zoom = 0;
                                        bookmarks[i].PageDisplay = "XYZ";
                                        bookmarksflag = 2;
                                    }
                                }
                            }
                            if (bookmarksflag == 2)
                            {
                                pdfEditor.DeleteBookmarks();
                                for (int bk = 0; bk < bookmarks.Count; bk++)
                                {
                                    if (bookmarks[bk].Level == 1)
                                        pdfEditor.CreateBookmarks(bookmarks[bk]);
                                }
                                pdfEditor.Save(sourcePath);
                            }
                            //pdfEditor.Dispose();
                            //pdfEditor.Close();
                            document = new Document(sourcePath);
                        }

                        for (int i = 1; i <= document.Pages.Count; i++)
                        {
                            Page page = document.Pages[i];
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                            page.Accept(selector);
                            IList<Annotation> list = selector.Selected;
                            foreach (LinkAnnotation a in list)
                            {
                                try
                                {
                                    if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                    {
                                        linksflag = 1;
                                        XYZExplicitDestination xyz = ((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination as XYZExplicitDestination;
                                        if (xyz == null)
                                        {
                                            int pageNo = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                            //XYZExplicitDestination xyznew = new XYZExplicitDestination(pageNo, 0, 0, 0.0);
                                            XYZExplicitDestination xyznew = new XYZExplicitDestination(pageNo, 0, page.PageInfo.Height, 0.0);
                                            GoToAction gta = new GoToAction(xyznew);
                                            a.Action = gta;
                                            linksflag = 2;
                                            linksFailedFlag = true;
                                        }
                                        else if ((xyz).Zoom > 0)
                                        {
                                            //XYZExplicitDestination dest = new XYZExplicitDestination(xyz.Page.Number, 0, 0, 0.0);
                                            XYZExplicitDestination xyznew = new XYZExplicitDestination(xyz.Page.Number, 0, page.PageInfo.Height, 0.0);
                                            GoToAction gta = new GoToAction(xyznew);
                                            a.Action = gta;
                                            linksflag = 2;
                                            linksFailedFlag = true;
                                        }
                                    }
                                    else if (a.Destination!= null && (a.Destination).ToString() != "")
                                    {
                                        XYZExplicitDestination xyz = a.Destination as XYZExplicitDestination;
                                        if (xyz == null || (xyz).Zoom > 0)
                                        {                                           
                                            linksflag = 2;
                                            linksFailedFlag = true;
                                        }
                                        else if ((xyz).Zoom == 0)
                                        {
                                            linksflag = 1;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (ex.Message == "Unable to cast object of type 'Aspose.Pdf.Annotations.NamedDestination' to type 'Aspose.Pdf.Annotations.ExplicitDestination'.")
                                    {
                                        IsNamedDestinationsExisted = true;
                                    }
                                }
                            }
                            page.FreeMemory();
                        }
                        if (IsNamedDestinationsExisted && linksFailedFlag && bookmarksflag == 2)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Inherit zoom is set for the bookmarks and links except NamedDestinations which cannot be fixed";
                        }
                        else if (IsNamedDestinationsExisted && linksFailedFlag && bookmarksflag == 1)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Inherit zoom is set for the links except NamedDestinations which cannot be fixed and bookmarks are already set to inherit zoom";
                        }
                        else if (IsNamedDestinationsExisted && linksFailedFlag && bookmarksflag == 0)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Inherit zoom is set for the links except NamedDestinations which cannot be fixed and no bookmarks exist in document";
                        }
                        else if (IsNamedDestinationsExisted && linksflag == 1 && bookmarksflag == 2)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Found NamedDestinations which cannot be fixed and bookmarks are set to inherit zoom";
                        }
                        else if (IsNamedDestinationsExisted && linksflag == 1 && bookmarksflag == 1)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Found NamedDestinations which cannot be fixed and bookmarks are already set to inherit zoom";
                        }
                        else if (IsNamedDestinationsExisted && linksflag == 1 && bookmarksflag == 0)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Found NamedDestinations which cannot be fixed and no bookmarks exist in document";
                        }
                        else if (linksFailedFlag && bookmarksflag == 2)
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Links and bookmarks are set to inherit zoom";
                        }
                        else if (linksFailedFlag && bookmarksflag == 1)
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Links are set to inherit zoom and bookmarks are already set to inherit zoom";
                        }
                        else if (linksFailedFlag && bookmarksflag == 0)
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Links are set to inherit zoom and no bookmarks exist in document";
                        }
                        else if (linksflag == 1 && bookmarksflag == 2)
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Links are already set to inherit zoom and bookmarks are set to inherit zoom";
                        }
                        else if (linksflag == 1 && bookmarksflag == 1)
                        {
                            rObj.Is_Fixed = 1;
                            //rObj.QC_Result = "Passed";
                            //rObj.Comments = "Links and bookmarks are already set to inherit zoom";
                            rObj.Comments = rObj.Comments + ", . Fixed";
                        }
                        else if (linksflag == 1 && bookmarksflag == 0)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Links are already set to inherit zoom and no bookmarks exist in document";
                        }
                        else if (linksflag == 0 && bookmarksflag == 2)
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Bookmarks are set to inherit zoom and no links exist in document";
                        }
                        else if (linksflag == 0 && bookmarksflag == 1)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are already set to inherit zoom and no links exist in document";
                        }
                        else if (linksflag == 0 && bookmarksflag == 1)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are not having destinations and no links exist in document";
                        }
                        else if (linksFailedFlag && bookmarksflag == 1)
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Links are set to inherit zoom and bookmarks are not having destinations";
                        }
                        else if (linksflag == 0 && bookmarksflag == 0)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "No links and bookmarks exist in document";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "There are no pages in the document";
                    }
                    document.Save(sourcePath);
                    //document.Dispose();
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
                    //else
                    //{
                    //    GoToAction action = new GoToAction();
                    //    action.Destination = new XYZExplicitDestination(1, 0.0, document.Pages[1].Rect.Height, 0.0);
                    //    document.OpenAction = action;
                    //}
                }
                catch
                {
                }

                document.Save(sourcePath);                
                System.IO.File.Copy(sourcePath, destPath, true);
                //document.Dispose();
            }
            catch (Exception ee)
            {
                ErrorLogger.Error(ee);
            }
        }

        public void RemoveRedundantBookmarksFix(RegOpsQC rObj, string path,Document document)
        {
            try
            {
                bool redundantExisted = false;
                string redundantBookmarks = string.Empty;
                List<string> bookmarkName = new List<string>();
                rObj.FIX_START_TIME = DateTime.Now;
                sourcePath = path + "//" + rObj.File_Name;

                //Document document = new Document(sourcePath);
                if (document.Pages.Count != 0)
                {
                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(sourcePath);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                    //bool flag = true;
                    if (bookmarks.Count > 0)
                    {
                        do
                        {
                            redundantExisted = false;
                            for (int i = 0; i < bookmarks.Count; i++)
                            {
                                if (!bookmarkName.Contains(bookmarks[i].Title + "_" + bookmarks[i].PageNumber))
                                {
                                    bookmarkName.Add(bookmarks[i].Title + "_" + bookmarks[i].PageNumber);
                                }
                                else
                                {
                                    redundantExisted = true;
                                    bookmarks.RemoveAt(i);
                                    bookmarkName.Clear();
                                    break;
                                }
                            }
                        }
                        while (redundantExisted == true);

                        Bookmark bk = null;
                        Bookmark BkLevel2 = null;
                        Bookmark BkLevel3 = null;
                        if (bookmarks.Count > 0)
                        {
                            bookmarkEditor.DeleteBookmarks();
                            for (int i = 0; i < bookmarks.Count; i++)
                            {
                                if (bk == null && bookmarks[i].Level == 1)
                                {
                                    bk = SetBookmarkPropertiesForRedundant(bookmarks[i]);
                                }
                                else if (bk != null && bookmarks[i].Level == 2)
                                {
                                    BkLevel2 = SetBookmarkPropertiesForRedundant(bookmarks[i]);
                                    bk.ChildItems.Add(BkLevel2);
                                }
                                else if (bk != null && bookmarks[i].Level == 3)
                                {
                                    if (bk.ChildItems.Count > 0)
                                    {
                                        BkLevel3 = SetBookmarkPropertiesForRedundant(bookmarks[i]);
                                        bk.ChildItems[bk.ChildItems.Count - 1].ChildItems.Add(BkLevel3);
                                    }
                                }
                                else if (bk != null && bookmarks[i].Level == 3)
                                {
                                    if (bk.ChildItems[bk.ChildItems.Count - 1].ChildItems.Count > 0)
                                    {
                                        BkLevel3 = SetBookmarkPropertiesForRedundant(bookmarks[i]);
                                        bk.ChildItems[bk.ChildItems.Count - 1].ChildItems[bk.ChildItems[bk.ChildItems.Count - 1].ChildItems.Count - 1].ChildItems.Add(BkLevel3);
                                    }
                                }
                                else if (bk != null && bookmarks[i].Level == 1)
                                {
                                    bookmarkEditor.CreateBookmarks(bk);
                                    bk = SetBookmarkPropertiesForRedundant(bookmarks[i]);
                                }

                            }
                            if (bk != null)
                            {
                                bookmarkEditor.CreateBookmarks(bk);
                            }
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = rObj.Comments + " .These bookmarks are removed from the document";
                            bookmarkEditor.Save(sourcePath);
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "No bookmarks exist in the document";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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
        public Bookmark SetBookmarkPropertiesForRedundant(Bookmark bookmark)
        {
            Bookmark bookmarkLeaf = null;
            try
            {
                bookmarkLeaf = new Bookmark();
                bookmarkLeaf.PageNumber = bookmark.PageNumber;
                bookmarkLeaf.Action = "GoTo";
                bookmarkLeaf.Title = bookmark.Title;
                bookmarkLeaf.PageDisplay = "XYZ";
                bookmarkLeaf.PageDisplay_Left = bookmark.PageDisplay_Left;
                bookmarkLeaf.PageDisplay_Top = bookmark.PageDisplay_Top;
                bookmarkLeaf.PageDisplay_Zoom = 0;
                bookmarkLeaf.Level = bookmark.Level;
                return bookmarkLeaf;
            }
            catch
            {
                return bookmarkLeaf;
            }
        }

        public Heading SetPropertiesForTOCItemsNew(RegOpsQC rObj, int Level, List<RegOpsQC> chLst)
        {
            Heading textFrg = new Heading(Level);
            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                chLst = chLst.Where(x => x.Check_Name.Contains("Level" + Level + " - Indentation")).ToList();
                //applying styles
                for (int fs = 0; fs < chLst.Count; fs++)
                {
                    if (chLst[fs].Check_Name == "Level" + Level + " - Indentation" && chLst[fs].Check_Type == 1)
                        textFrg.Margin.Left = float.Parse(chLst[fs].Check_Parameter);
                }
                return textFrg;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error(ee);
                return null;
            }
        }

        /// <summary>
        /// Properties fields should be blank except title
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="checkString"></param>
        /// <param name="destPath"></param>
        /// <param name="checkType"></param>
        public void PDFFile_PropertiesExceptTitleCheck(RegOpsQC rObj, string path, string checkString, double checkType,Document document)
        {
            string res = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);
                string title = document.Info.Title;
                if ((document.Info.Author != null && document.Info.Author != "") || (document.Info.Subject != null && document.Info.Subject != "") || (document.Info.Keywords != null && document.Info.Keywords != ""))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Property fields are not blank";
                }
                else if ((document.Info.Title != null && document.Info.Title != "") && (document.Info.Author == null || document.Info.Author == "") &&
                    (document.Info.Subject == null || document.Info.Subject == "") && (document.Info.Keywords == null || document.Info.Keywords == ""))
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Property fields are blank except title";
                }
                else if ((document.Info.Title == null || document.Info.Title == "") && (document.Info.Author == null || document.Info.Author == "") &&
                    (document.Info.Subject == null || document.Info.Subject == "") && (document.Info.Keywords == null || document.Info.Keywords == ""))
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All property fields are blank including title";
                }
                //document.Dispose();
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
        /// Properties fields should be blank except title - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="checkString"></param>
        /// <param name="destPath"></param>
        /// <param name="checkType"></param>
        public void PDFFile_PropertiesExceptTitleFix(RegOpsQC rObj, string path, string checkString, double checkType,Document document)
        {
            string res = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);
                string title = document.Info.Title;
                document.RemoveMetadata();
                document.Info.Author = string.Empty;
                document.Info.Subject = string.Empty;
                document.Info.Keywords = string.Empty;
                document.Info.Title = string.Empty;
                document.Info.Title = title;
                //document.Save(sourcePath);
                //document.Dispose();
                //rObj.QC_Result = "Fixed";
                rObj.Is_Fixed = 1;
                rObj.Comments = "Default properties set to blank except title";
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
        /// Create TOC must for above 5 page and more - Check latest
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="chLst"></param>
        public void CreateTOCFromBookmarks(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document myDocument)
        {
            bool isTOCExisted = false;
            bool isLOTExisted = false;
            bool isLOFExisted = false;
            bool isScannedDoc = true;
            int lotbookmrkflag = 0;
            int lofbookmrkflag = 0;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
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
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                //Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                               
               // Document myDocument = new Document(sourcePath);
                Regex rx_toc = new Regex(@"(TABLE OF CONTENTS\s?\r\n|TABLE OF CONTENTS\s+\r\n|TABLE OF CONTENT\s?\r\n|TABLE OF CONTENT\s+\r\n)", RegexOptions.IgnoreCase);
                Regex rx_lot = new Regex(@"(LIST OF TABLES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex rx_lof = new Regex(@"(LIST OF FIGURES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex regextbl = new Regex(@"(Table\s\d)", RegexOptions.IgnoreCase);
                Regex regexfig = new Regex(@"(Figure\s\d)", RegexOptions.IgnoreCase);

                Dictionary<int, Bookmark> dictLOT = new Dictionary<int, Bookmark>();
                Dictionary<int, Bookmark> dictLOF = new Dictionary<int, Bookmark>();
                if (myDocument.Pages.Count >= 5)
                {
                    if (bookmarks.Count > 0)
                    {
                        //Below code is to check whether LOT and LOF are existed in the bookmarks or not.
                        for (int i = 0; i < bookmarks.Count(); i++)
                        {
                            //if (Regex.IsMatch(bookmarks[i].Title.ToUpper(), @"(TABLE OF CONTENTS\s?|TABLE OF CONTENT\s?)",RegexOptions.IgnoreCase) && bookmarks[i].PageNumber > 0)
                            //{
                            //    isTOCExisted = true;
                            //    rObj.isTOCExisted = true;
                            //}
                            if ((bookmarks[i].Title.ToUpper() == "LIST OF TABLES" && bookmarks[i].ChildItems.Count > 0)|| regextbl.IsMatch(bookmarks[i].Title))
                            {
                                lotbookmrkflag = 1;
                            }
                            if ((bookmarks[i].Title.ToUpper() == "LIST OF FIGURES" && bookmarks[i].ChildItems.Count > 0)|| regexfig.IsMatch(bookmarks[i].Title))
                            {
                                lofbookmrkflag = 1;
                            }
                            if (lotbookmrkflag == 1 && lofbookmrkflag == 1)
                                break;
                        }
                        //Checking whether TOC existed in the current document by checking each page upto 10 pages.
                        Document currentDocToCheckTOC = new Document(sourcePath);
                        System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"(Table of Content)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        TextFragmentAbsorber textbsorber = new TextFragmentAbsorber(regex);
                        Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                        textbsorber.TextSearchOptions = textSearchOptions;
                        for (int i = 1; i <= currentDocToCheckTOC.Pages.Count; i++)
                        {
                            using (MemoryStream textStream = new MemoryStream())
                            {
                                // Create text device
                                TextDevice textDevice = new TextDevice();
                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                textDevice.ExtractionOptions = textExtOptions;
                                textDevice.Process(currentDocToCheckTOC.Pages[i], textStream);
                                // Close memory stream
                                textStream.Close();
                                // Get text from memory stream
                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                if (extractedText != "")
                                {

                                    isScannedDoc = false;
                                    //Matching regular expression to check the TOC existence
                                    if (rx_toc.IsMatch(extractedText) && rObj.isTOCExisted == false)
                                    {
                                        isTOCExisted = true;
                                        rObj.isTOCExisted = true;
                                        //isTOCLinksexisted = true;
                                        //break;                                   
                                    }
                                    //Matching regular expression to check the LOT existence
                                    if (rx_lot.IsMatch(extractedText))
                                    {
                                        isLOTExisted = true;
                                        rObj.isLOTExisted = true;
                                        //isTOCLinksexisted = true;
                                        //break;
                                    }
                                    //Matching regular expression to check the LOF existence
                                    if (rx_lof.IsMatch(extractedText))
                                    {
                                        isLOFExisted = true;
                                        rObj.isLOFExisted = true;
                                        //isTOCLinksexisted = true;
                                        //break;
                                    }
                                }
                            }
                            if (i == 10)
                                break;
                            else if (isTOCExisted && isLOFExisted && isLOTExisted)
                                break;
                        }
                        currentDocToCheckTOC.Dispose();
                        if (isScannedDoc)
                        {
                            rObj.Comments = "Cannot check TOC,LOT and LOF for scanned documents";
                            rObj.QC_Result = "Failed";
                        }
                        else
                        {
                            if (isTOCExisted && isLOFExisted && isLOTExisted)
                            {
                                //rObj.Comments = "TOC, LOT and LOF are existed in the document";
                                rObj.QC_Result = "Passed";
                            }
                            else if (isTOCExisted && isLOTExisted && !isLOFExisted && lofbookmrkflag == 1)
                            {
                                rObj.Comments = "TOC,LOT exist but LOF does not exist";
                                rObj.QC_Result = "Failed";
                            }
                            else if (isTOCExisted && !isLOTExisted && !isLOFExisted && lotbookmrkflag == 1 && lofbookmrkflag == 1)
                            {
                                rObj.Comments = "TOC exist but LOT and LOF does not exist";
                                rObj.QC_Result = "Failed";
                            }                                               
                            else if (!isTOCExisted && isLOTExisted && isLOFExisted)
                            {
                                rObj.Comments = "LOT and LOF exist but not TOC";
                                rObj.QC_Result = "Failed";
                            }
                            else if (!isTOCExisted && !isLOTExisted && isLOFExisted && lotbookmrkflag == 1)
                            {
                                rObj.Comments = "LOF exists but not TOC and LOT";
                                rObj.QC_Result = "Failed";
                            }
                            else if (!isTOCExisted && !isLOTExisted && !isLOFExisted && lotbookmrkflag == 1 && lofbookmrkflag == 1)
                            {
                                rObj.Comments = "TOC, LOT and LOF do not exist";
                                rObj.QC_Result = "Failed";
                            }
                            else if (isTOCExisted && !isLOTExisted && !isLOFExisted && lotbookmrkflag == 0 && lofbookmrkflag == 0)
                            {
                                //rObj.Comments = "TOC existed in the document";
                                rObj.QC_Result = "Passed";
                            }
                            else if (isTOCExisted && !isLOTExisted && isLOFExisted && lotbookmrkflag == 0)
                            {
                                //rObj.Comments = "TOC existed in the document";
                                rObj.QC_Result = "Passed";
                            }
                            else if (isTOCExisted && isLOTExisted && !isLOFExisted && lofbookmrkflag == 0)
                            {
                                //rObj.Comments = "TOC existed in the document";
                                rObj.QC_Result = "Passed";
                            }
                            else if (!isTOCExisted && !isLOTExisted && !isLOFExisted && lotbookmrkflag == 1 && lofbookmrkflag == 0)
                            {
                                rObj.Comments = "TOC, LOT do not exist";
                                rObj.QC_Result = "Failed";
                            }
                            else if (!isTOCExisted && !isLOTExisted && !isLOFExisted && lotbookmrkflag == 0 && lofbookmrkflag == 1)
                            {
                                rObj.Comments = "TOC, LOF do not exist";
                                rObj.QC_Result = "Failed";
                            }
                            else if (isTOCExisted && !isLOTExisted && !isLOFExisted && (lotbookmrkflag == 1 && lofbookmrkflag == 0 || lotbookmrkflag == 0 && lofbookmrkflag == 1) || lotbookmrkflag == 1 && lofbookmrkflag == 1)
                            {
                                rObj.Comments = "TOC exist but LOT LOF does not exist";
                                rObj.QC_Result = "Failed";
                            }
                            else
                            {
                                rObj.Comments = "TOC does not exist";
                                rObj.QC_Result = "Failed";
                            }
                        }
                    }
                    else if (bookmarks.Count == 0)
                    {
                        rObj.Comments = "No bookmarks existed in document. So TOC cannot be created";
                        rObj.QC_Result = "Failed";
                    }
                }
                else if (myDocument.Pages.Count < 5)
                {
                    //rObj.Comments = "Document has less than 5 pages.";
                    rObj.QC_Result = "Passed";
                }
                //bookmarkEditor.Dispose();
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
        /// Create TOC must for above 5 page and more - fix before code optimization. this method is commented on 14-Dec-2020.
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="chLst"></param>
        //public void CreateTOCFromBookmarksFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst)
        //{            
        //    sourcePath = path + "//" + rObj.File_Name;

        //    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        //    List<string> bookmarkkName = new List<string>();
        //    //Open PDF file
        //    bookmarkEditor.BindPdf(sourcePath);
        //    //Extracting bookmarks from the source doc.
        //    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
        //    bookmarkEditor.Close();

        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    try
        //    {
        //        //Checking for the TOC only check or fix
        //        bool fixornot = false;
        //        chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
        //        for (int i = 0; i < chLst.Count; i++)
        //        {
        //            chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
        //            chLst[i].JID = rObj.JID;
        //            chLst[i].Job_ID = rObj.Job_ID;
        //            chLst[i].Folder_Name = rObj.Folder_Name;
        //            chLst[i].File_Name = rObj.File_Name;
        //            chLst[i].Created_ID = rObj.Created_ID;

        //            if (chLst[i].Check_Type == 1)
        //            {
        //                fixornot = true;
        //                //break;                       
        //            }
        //        }
        //        if (fixornot && bookmarks.Count > 0)
        //        {
        //            //Loading source document
        //            Document myDocument = new Document(sourcePath);
        //            int initialPageCount = myDocument.Pages.Count();

        //            //Getting page size of the original document.
        //            PdfPageEditor pageEditor = new PdfPageEditor();
        //            pageEditor.BindPdf(sourcePath);
        //            PageSize originalPG = pageEditor.GetPageSize(1);
        //            pageEditor.Close();

        //            TocInfo tocInfo = new TocInfo();
        //            Aspose.Pdf.Page tocPage = null;

        //            TextFragment titleFrag = new TextFragment();

        //            bool tocTitleFlag = false;
        //            bool lotFlag = false;
        //            bool lofFlag = false;
        //            bool isLOTExisted = false;
        //            bool isLOFExisted = false;
        //            Guid guid = Guid.NewGuid();

        //            int TOCPgCnt = 0;
        //            int LOTPgCnt = 0;
        //            int LOFPgCnt = 0;
        //            bool createtoc = false;
        //            int tocbkexist = 0;
        //            bool istablebkmrksexist = false;
        //            bool isfigurebkmrksexist = false;
        //            int starttbl = -1, endtbl = -1;
        //            int startfig = -1, endfig = -1;
        //            int lotbkcount = 0, lofbkcount = 0;
        //            Regex regextbl = new Regex(@"(Table\s\d)");
        //            Regex regexfig = new Regex(@"(Figure\s\d)");
        //            //If any duplicate TOC bookmark existed removing it                    
        //            for (int i = 0; i < bookmarks.Count(); i++)
        //            {
        //                if (bookmarks[i].Title.ToUpper() == "TABLE OF CONTENTS" && bookmarks[i].ChildItems.Count == 0)
        //                {
        //                    tocbkexist = 1;
        //                }
        //                if (bookmarks[i].Title.ToUpper() == "LIST OF TABLES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    isLOTExisted = true;
        //                }
        //                if (bookmarks[i].Title.ToUpper() == "LIST OF FIGURES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    isLOFExisted = true;
        //                }
        //                if (bookmarks[i].Title.ToUpper() != "TABLE OF CONTENTS" && bookmarks[i].Title.ToUpper() != "LIST OF TABLES" && bookmarks[i].Title.ToUpper() != "LIST OF FIGURES" && bookmarks[i].Level == 1 && tocbkexist == 0)
        //                {
        //                    createtoc = true;
        //                }
        //                if (!isLOTExisted && regextbl.IsMatch(bookmarks[i].Title))
        //                {
        //                    if (starttbl==-1)
        //                    {
        //                        starttbl = i;
        //                    }
        //                    else 
        //                    {
        //                        if (i - endtbl == 1)
        //                        {
        //                            istablebkmrksexist = true;
        //                        }
        //                        else
        //                        {
        //                            istablebkmrksexist = false;
        //                        }
        //                    }
        //                    endtbl = i;
        //                    lotbkcount++;
        //                }
        //                if (!isLOFExisted && regexfig.IsMatch(bookmarks[i].Title))
        //                {
        //                    if (startfig == -1)
        //                    {
        //                        startfig = i;
        //                    }
        //                    else
        //                    {
        //                        if (i- endfig==1)
        //                        {
        //                            isfigurebkmrksexist = true;
        //                        }
        //                        else
        //                        {
        //                            isfigurebkmrksexist = false;
        //                        }
        //                    }
        //                    endfig = i;
        //                    lofbkcount++;
        //                }
        //            }
        //            if (istablebkmrksexist)
        //            {
        //                isLOTExisted = true;
        //            }
        //            if (isfigurebkmrksexist)
        //            {
        //                isLOFExisted = true;
        //            }
        //            int lotCount = 0;
        //            int lofCount = 0;
        //            int flagCount = 0;
        //            int LOTPagePosition = 0;
        //            int LOFPagePosition = 0;
        //            bool isLOTPageIdentified = false;
        //            bool isLOFPageIdentified = false;
        //            bool isTOCCompleted = false;
        //            bool skiplot = false;
        //            bool skiplof = false;
        //            string InprogressFlag = string.Empty;
        //            int insertedpgs = 0;
        //            for (int i = 0; i < bookmarks.Count; i++)
        //            {
        //                string title = bookmarks[i].Title;
        //                if ((InprogressFlag == "TOC"||( InprogressFlag=="")) && title.ToUpper() == "LIST OF TABLES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    int count = i + bookmarks[i].ChildItems.Count + 1;
        //                    if (isTOCCompleted == false)
        //                    {
        //                        if (bookmarks.Count > count && !rObj.isTOCExisted)
        //                            i = i + bookmarks[i].ChildItems.Count + 1;
        //                        else
        //                            isTOCCompleted = true;
        //                    }                           

        //                    title = bookmarks[i].Title;
        //                }
        //                if ((InprogressFlag == "TOC"||(InprogressFlag == "")) && title.ToUpper() == "LIST OF FIGURES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    int count  = i + bookmarks[i].ChildItems.Count + 1;
        //                    if(isTOCCompleted==false)
        //                    {
        //                        if (bookmarks.Count > count && !rObj.isTOCExisted)
        //                            i = i + bookmarks[i].ChildItems.Count + 1;
        //                        else
        //                            isTOCCompleted = true;
        //                    }                                
        //                    title = bookmarks[i].Title;

        //                }
        //                if (InprogressFlag == "LOF" && title.ToUpper() == "LIST OF TABLES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    int count = i + bookmarks[i].ChildItems.Count + 1;
        //                    if (isTOCCompleted == false)
        //                    {
        //                        if (bookmarks.Count > count && !rObj.isTOCExisted)
        //                            i = i + bookmarks[i].ChildItems.Count + 1;
        //                        else
        //                            isTOCCompleted = true;
        //                    }

        //                    title = bookmarks[i].Title;
        //                }
        //                if (InprogressFlag == "TOC" && isLOTExisted && istablebkmrksexist && !skiplot)
        //                {
        //                    i = endtbl + 1;
        //                    title = bookmarks[i].Title;
        //                    skiplot = true;
        //                }
        //                if (InprogressFlag == "TOC" && isLOFExisted && isfigurebkmrksexist && !skiplof)
        //                {
        //                    i = endfig + 1;
        //                    title = bookmarks[i].Title;
        //                    skiplof = true;
        //                }
        //                title = bookmarks[i].Title;
        //                //Checking whether toc existed or not
        //                if ((tocTitleFlag == false && isTOCCompleted == false && bookmarks[i].Level == 1 && (title.ToUpper() != "LIST OF TABLES") && title.ToUpper() != "LIST OF FIGURES" && rObj.isTOCExisted == false) || (createtoc == true && tocTitleFlag == false && rObj.isTOCExisted == false && isTOCCompleted == false))
        //                {
        //                    InprogressFlag = "TOC";
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "TABLE OF CONTENTS";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;

        //                    tocPage = myDocument.Pages.Insert(1);

        //                    //Set Page size for TOC pages
        //                    if (originalPG.Width > originalPG.Height)
        //                        myDocument.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
        //                    else
        //                        myDocument.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);

        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;

        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    tocTitleFlag = true;
        //                    //rObj.isTOCExisted = true;
        //                    if (createtoc == true && InprogressFlag == "TOC")
        //                    {
        //                        i = -1;
        //                    }
        //                }
        //                //Checking whether TOC and LOT existed or created
        //                else if ((title.ToUpper() == "LIST OF TABLES"|| (istablebkmrksexist && isLOTExisted)) && (isTOCCompleted || rObj.isTOCExisted == true) && lotFlag == false && rObj.isLOTExisted == false)
        //                {
        //                    InprogressFlag = "LOT";
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF TABLES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;

        //                    myDocument.Save(sourcePath1 + guid + rObj.File_Name);
        //                    myDocument.Dispose();
        //                    myDocument = new Document(sourcePath1 + guid + rObj.File_Name);

        //                    if (isTOCCompleted || rObj.isTOCExisted == true)
        //                    {
        //                        TOCPgCnt = 0;

        //                        //Need to verify where toc page ended
        //                        for (int k = 1; k <= myDocument.Pages.Count; k++)
        //                        {
        //                            using (MemoryStream textStream = new MemoryStream())
        //                            {
        //                                // Create text device
        //                                TextDevice textDevice = new TextDevice();
        //                                // Set text extraction options - set text extraction mode (Raw or Pure)
        //                                Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
        //                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
        //                                textDevice.ExtractionOptions = textExtOptions;
        //                                textDevice.Process(myDocument.Pages[k], textStream);
        //                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
        //                                // Close memory stream
        //                                textStream.Close();
        //                                // Get text from memory stream
        //                                Regex regex = null;
        //                                if (extractedText.ToUpper().Contains("TABLE OF CONTENT"))
        //                                {
        //                                    for (int T = k; T <= myDocument.Pages.Count; T++)
        //                                    {
        //                                        using (MemoryStream textStreamTemp = new MemoryStream())
        //                                        {
        //                                            // Create text device
        //                                            textDevice = new TextDevice();
        //                                            // Set text extraction options - set text extraction mode (Raw or Pure)
        //                                            textExtOptions = new
        //                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
        //                                            textDevice.ExtractionOptions = textExtOptions;
        //                                            textDevice.Process(myDocument.Pages[T], textStreamTemp);
        //                                            extractedText = Encoding.Unicode.GetString(textStreamTemp.ToArray());
        //                                            // Close memory stream  
        //                                            textStreamTemp.Close();
        //                                            regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //                                            if ((extractedText.ToUpper().Contains("LIST OF TABLES") || extractedText.ToUpper().Contains("LIST OF FIGURES") || !regex.IsMatch(extractedText)) && !extractedText.ToUpper().Contains("TABLE OF CONTENT"))
        //                                            {
        //                                                LOTPagePosition = T;
        //                                                //Inserting a new page for LOT after TOC.
        //                                                tocPage = myDocument.Pages.Insert(T);
        //                                                //Set Page size for TOC pages
        //                                                if (originalPG.Width > originalPG.Height)
        //                                                    myDocument.Pages[T].SetPageSize(originalPG.Height, originalPG.Width);
        //                                                else
        //                                                    myDocument.Pages[T].SetPageSize(originalPG.Width, originalPG.Height);

        //                                                isLOTPageIdentified = true;
        //                                                break;
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                            if (isLOTPageIdentified)
        //                                break;
        //                        }
        //                    }
        //                    else if (tocTitleFlag)
        //                    {
        //                        TOCPgCnt = myDocument.Pages.Count - initialPageCount;
        //                        tocPage = myDocument.Pages.Insert(TOCPgCnt + 1);
        //                        LOTPagePosition = TOCPgCnt + 1;

        //                        //Set Page size for TOC pages
        //                        if (originalPG.Width > originalPG.Height)
        //                            myDocument.Pages[LOTPagePosition].SetPageSize(originalPG.Height, originalPG.Width);
        //                        else
        //                            myDocument.Pages[LOTPagePosition].SetPageSize(originalPG.Width, originalPG.Height);
        //                    }
        //                    if (isLOTPageIdentified)
        //                    {
        //                        tocTitleFlag = true;
        //                    }

        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;

        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    lotFlag = true;
        //                    if (istablebkmrksexist && isLOTExisted)
        //                        lotCount = lotbkcount;
        //                    else
        //                        lotCount = bookmarks[i].ChildItems.Count;
        //                    lofCount = 0;
        //                    flagCount = 0;
        //                }
        //                else if ((title.ToUpper() == "LIST OF FIGURES"|| (isfigurebkmrksexist&& isLOFExisted)) && rObj.isLOFExisted == false && isLOTExisted == false && isLOFExisted == true && (isTOCCompleted == true || rObj.isTOCExisted == true))
        //                {
        //                    InprogressFlag = "LOF";
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF FIGURES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;

        //                    myDocument.Save(sourcePath1 + guid + rObj.File_Name);
        //                    myDocument.Dispose();

        //                    myDocument = new Document(sourcePath1 + guid + rObj.File_Name);
        //                    if (isTOCCompleted || rObj.isTOCExisted == true)
        //                    {
        //                        //Need to verify where LOT page ended
        //                        for (int k = 1; k <= myDocument.Pages.Count; k++)
        //                        {
        //                            using (MemoryStream textStream = new MemoryStream())
        //                            {
        //                                // Create text device
        //                                TextDevice textDevice = new TextDevice();
        //                                // Set text extraction options - set text extraction mode (Raw or Pure)
        //                                Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
        //                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
        //                                textDevice.ExtractionOptions = textExtOptions;
        //                                textDevice.Process(myDocument.Pages[k], textStream);
        //                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
        //                                // Close memory stream
        //                                textStream.Close();
        //                                // Get text from memory stream
        //                                Regex regex = null;
        //                                if (extractedText.ToUpper().Contains("TABLE OF CONTENT"))
        //                                {
        //                                    for (int T = k; T <= myDocument.Pages.Count; T++)
        //                                    {
        //                                        using (MemoryStream textStreamTemp = new MemoryStream())
        //                                        {
        //                                            // Create text device
        //                                            textDevice = new TextDevice();
        //                                            // Set text extraction options - set text extraction mode (Raw or Pure)
        //                                            textExtOptions = new
        //                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
        //                                            textDevice.ExtractionOptions = textExtOptions;
        //                                            textDevice.Process(myDocument.Pages[T], textStreamTemp);
        //                                            extractedText = Encoding.Unicode.GetString(textStreamTemp.ToArray());
        //                                            // Close memory stream  
        //                                            textStreamTemp.Close();
        //                                            regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //                                            if (extractedText.ToUpper().Contains("LIST OF FIGURES") || !regex.IsMatch(extractedText))
        //                                            {
        //                                                isLOFPageIdentified = true;
        //                                                //Inserting a new page for LOF after TOC and LOT
        //                                                tocPage = myDocument.Pages.Insert(T);
        //                                                LOFPagePosition = T;

        //                                                //Set Page size for TOC pages
        //                                                if (originalPG.Width > originalPG.Height)
        //                                                    myDocument.Pages[T].SetPageSize(originalPG.Height, originalPG.Width);
        //                                                else
        //                                                    myDocument.Pages[T].SetPageSize(originalPG.Width, originalPG.Height);

        //                                                break;
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                            if (isLOFPageIdentified)
        //                                break;
        //                        }
        //                    }
        //                    if (isLOFPageIdentified)
        //                    {
        //                        tocTitleFlag = true;
        //                    }

        //                    //Title will be added here
        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;

        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    if (isfigurebkmrksexist && isLOFExisted)
        //                        lofCount = lofbkcount;
        //                    else
        //                        lofCount = bookmarks[i].ChildItems.Count;
        //                    lofFlag = true;
        //                    flagCount = 0;
        //                    lotCount = 0;
        //                }
        //                //Checking whether TOC, LOT and LOF existed or created
        //                else if ((title.ToUpper() == "LIST OF FIGURES" || (isfigurebkmrksexist && isLOFExisted)) && ((lotFlag == true && isLOTExisted) || (lotFlag == false && !isLOTExisted) || (rObj.isTOCExisted && rObj.isLOTExisted && rObj.isLOFExisted == false)) && rObj.isLOFExisted == false)
        //                {
        //                    InprogressFlag = "LOF";
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF FIGURES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;

        //                    myDocument.Save(sourcePath1 + guid + rObj.File_Name);
        //                    myDocument.Dispose();

        //                    myDocument = new Document(sourcePath1 + guid + rObj.File_Name);
        //                    if ((isLOTPageIdentified && LOTPagePosition != 0) || (rObj.isTOCExisted && rObj.isLOTExisted))
        //                    {

        //                        //Need to verify where LOT page ended
        //                        for (int k = 1; k <= myDocument.Pages.Count; k++)
        //                        {
        //                            using (MemoryStream textStream = new MemoryStream())
        //                            {
        //                                // Create text device
        //                                TextDevice textDevice = new TextDevice();
        //                                // Set text extraction options - set text extraction mode (Raw or Pure)
        //                                Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
        //                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
        //                                textDevice.ExtractionOptions = textExtOptions;
        //                                textDevice.Process(myDocument.Pages[k], textStream);
        //                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
        //                                // Close memory stream
        //                                textStream.Close();
        //                                // Get text from memory stream
        //                                Regex regex = null;
        //                                if (extractedText.ToUpper().Contains("LIST OF TABLES"))
        //                                {
        //                                    for (int T = k; T <= myDocument.Pages.Count; T++)
        //                                    {
        //                                        using (MemoryStream textStreamTemp = new MemoryStream())
        //                                        {
        //                                            // Create text device
        //                                            textDevice = new TextDevice();
        //                                            // Set text extraction options - set text extraction mode (Raw or Pure)
        //                                            textExtOptions = new
        //                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
        //                                            textDevice.ExtractionOptions = textExtOptions;
        //                                            textDevice.Process(myDocument.Pages[T], textStreamTemp);
        //                                            extractedText = Encoding.Unicode.GetString(textStreamTemp.ToArray());
        //                                            // Close memory stream  
        //                                            textStreamTemp.Close();
        //                                            regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //                                            if (extractedText.ToUpper().Contains("LIST OF FIGURES") || !regex.IsMatch(extractedText))
        //                                            {
        //                                                isLOFPageIdentified = true;
        //                                                //Inserting a new page for LOF after TOC and LOT
        //                                                tocPage = myDocument.Pages.Insert(T);
        //                                                LOFPagePosition = T;

        //                                                //Set Page size for TOC pages
        //                                                if (originalPG.Width > originalPG.Height)
        //                                                    myDocument.Pages[T].SetPageSize(originalPG.Height, originalPG.Width);
        //                                                else
        //                                                    myDocument.Pages[T].SetPageSize(originalPG.Width, originalPG.Height);

        //                                                break;
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                            if (isLOFPageIdentified)
        //                                break;
        //                        }
        //                    }
        //                    if (isLOFPageIdentified)
        //                    {
        //                        tocTitleFlag = true;
        //                    }

        //                    //Title will be added here
        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;

        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    if (isfigurebkmrksexist && isLOFExisted)
        //                        lofCount = lofbkcount;
        //                    else
        //                    lofCount = bookmarks[i].ChildItems.Count;
        //                    lofFlag = true;
        //                    flagCount = 0;
        //                    lotCount = 0;
        //                }
        //                //Creating bookmarks to TOC/LOT/LOF                                                
        //                if (tocTitleFlag == true && title.ToUpper() != "TABLE OF CONTENTS" && title.ToUpper() != "LIST OF TABLES" && title.ToUpper() != "LIST OF FIGURES" && i>=0 && bookmarks[i].Level <= 4)
        //                {
        //                    if (lotCount > 0)
        //                    {
        //                        flagCount++;
        //                    }
        //                    if (lofCount > 0)
        //                    {
        //                        flagCount++;
        //                    }
        //                    Aspose.Pdf.Heading heading2 = new Heading(1);
        //                    heading2 = SetPropertiesForTOCItemsNew(rObj, bookmarks[i].Level, chLst);
        //                    //TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst);

        //                    TextSegment segment2 = new TextSegment();
        //                    heading2.TocPage = tocPage;
        //                    segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst, title);
        //                    heading2.TextState.ForegroundColor = Color.Blue;
        //                    heading2.TextState.LineSpacing = 10;
        //                    insertedpgs = myDocument.Pages.Count - initialPageCount;
        //                    // Destination page
        //                    heading2.DestinationPage = myDocument.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber + insertedpgs)]; //myDocument.Pages[i + 2];                        
        //                    heading2.Top = myDocument.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber + insertedpgs)].Rect.Height;
        //                    //LocalHyperlink lhl = new LocalHyperlink();
        //                    //lhl.TargetPageNumber = (bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber) + tocPages;
        //                    //segment2.Hyperlink = lhl;

        //                    if (bookmarks[i].Level == 2)
        //                    {
        //                        heading2.Margin.Left = 20;
        //                    }
        //                    else if (bookmarks[i].Level == 3)
        //                    {
        //                        heading2.Margin.Left = 25;
        //                    }
        //                    else if (bookmarks[i].Level == 4 || bookmarks[i].Level > 4)
        //                    {
        //                        heading2.Margin.Left = 30;
        //                    }

        //                    heading2.Segments.Add(segment2);
        //                    tocPage.Paragraphs.Add(heading2);

        //                }

        //                if (lofCount > 0 && flagCount == lofCount && flagCount > 0)                        
        //                    break;                        
        //                if (lotCount > 0 && flagCount == lotCount && isLOFExisted == false)
        //                    break;
        //                //Checking TOC completed or not
        //                if (i == bookmarks.Count - 1 && InprogressFlag == "TOC")
        //                {
        //                    tocTitleFlag = true;
        //                    InprogressFlag = string.Empty;
        //                    isTOCCompleted = true;
        //                }
        //                //if (createtoc == true && InprogressFlag == "TOC" && isTOCCompleted == false)
        //                //{
        //                //    i = -1;
        //                //}
        //                if (isLOTExisted == false && isLOFExisted == true && isTOCCompleted == true && tocTitleFlag == true && i == bookmarks.Count() - 1 && rObj.isLOFExisted == false)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                    tocTitleFlag = false;
        //                    InprogressFlag = string.Empty;
        //                }
        //                //Checking TOC completed and whether LOT  need to be created or nnot
        //                if (lotFlag == false && isLOTExisted == true && tocTitleFlag == true && i == bookmarks.Count() - 1 && rObj.isLOTExisted == false)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                    tocTitleFlag = false;
        //                    InprogressFlag = string.Empty;
        //                }
        //                //Checking TOC and LOT completed and whether LOF  need to be created or nnot
        //                else if (lotCount > 0 && flagCount == lotCount && rObj.isLOFExisted == false && isLOFExisted == true)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                    tocTitleFlag = false;
        //                    InprogressFlag = string.Empty;
        //                }
        //                else if(lotCount > 0 && flagCount == lotCount && isLOFExisted)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                    tocTitleFlag = false;
        //                    InprogressFlag = string.Empty;
        //                }
        //                //Exit from the loop if all are created.                        
        //                else if (lofCount > 0 && flagCount == lofCount && flagCount > 0)
        //                    break;
        //            }

        //            guid = Guid.NewGuid();
        //            //Saving the document
        //            myDocument.Save(sourcePath1 + guid + rObj.File_Name);
        //            myDocument.Dispose();

        //            myDocument = new Document(sourcePath1 + guid + rObj.File_Name);
        //            //int PagesNeedToBeAdded = 0;
        //            int tocPages = myDocument.Pages.Count - initialPageCount;
        //            myDocument.Dispose();

        //            //if (LOTPgCnt > 0)
        //            //    PagesNeedToBeAdded = LOTPgCnt + 1;
        //            //else if (LOTPgCnt == 0)
        //            //    PagesNeedToBeAdded = TOCPgCnt + 1;
        //            //else if (LOTPgCnt == 0 && LOFPgCnt == 0)
        //            //    PagesNeedToBeAdded = myDocument.Pages.Count - initialPageCount;




        //            ////Again taking source file as input file
        //            //myDocument = new Document(sourcePath);
        //            //for (int n = 1; n <= tocPages; n++)
        //            //{
        //            //    myDocument.Pages.Add();
        //            //    //myDocument.Pages.Insert(n);
        //            //}
        //            //myDocument.Save(sourcePath);
        //            //myDocument.Dispose();

        //            //The below code is used to create toc in fuinal document
        //            Document myDocumentNew = new Document(sourcePath);
        //            tocInfo = new TocInfo();

        //            isLOTPageIdentified = false;
        //            isLOFPageIdentified = false;
        //            isTOCCompleted = false;
        //            InprogressFlag = string.Empty;
        //            tocTitleFlag = false;
        //            lotFlag = false;
        //            lofFlag = false;
        //            lotCount = 0;
        //            lofCount = 0;
        //            flagCount = 0;
        //            TextFragment titleFrag1 = new TextFragment();
        //            int initialPageNo = 0;
        //            int addpages = 0;
        //            bool skiplotbks = false;
        //            bool skiplofbks = false;
        //            for (int i = 0; i < bookmarks.Count; i++)
        //            {
        //                string title = bookmarks[i].Title;

        //                if ((InprogressFlag == "TOC" || (InprogressFlag == "")) && title.ToUpper() == "LIST OF TABLES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    int count = i + bookmarks[i].ChildItems.Count + 1;
        //                    if (isTOCCompleted == false)
        //                    {
        //                        if (bookmarks.Count > count && !rObj.isTOCExisted)
        //                            i = i + bookmarks[i].ChildItems.Count + 1;
        //                        else
        //                            isTOCCompleted = true;
        //                    }

        //                    title = bookmarks[i].Title;
        //                }
        //                if ((InprogressFlag == "TOC" || (InprogressFlag == "")) && title.ToUpper() == "LIST OF FIGURES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    int count = i + bookmarks[i].ChildItems.Count + 1;
        //                    if (isTOCCompleted == false)
        //                    {
        //                        if (bookmarks.Count > count && !rObj.isTOCExisted)
        //                            i = i + bookmarks[i].ChildItems.Count + 1;
        //                        else
        //                            isTOCCompleted = true;
        //                    }

        //                    title = bookmarks[i].Title;

        //                }
        //                if (InprogressFlag == "LOF" && title.ToUpper() == "LIST OF TABLES" && bookmarks[i].ChildItems.Count > 0)
        //                {
        //                    int count = i + bookmarks[i].ChildItems.Count + 1;
        //                    if (isTOCCompleted == false)
        //                    {
        //                        if (bookmarks.Count > count && !rObj.isTOCExisted)
        //                            i = i + bookmarks[i].ChildItems.Count + 1;
        //                        else
        //                            isTOCCompleted = true;
        //                    }

        //                    title = bookmarks[i].Title;
        //                }
        //                if (InprogressFlag == "TOC" && isLOTExisted && istablebkmrksexist && !skiplotbks)
        //                {
        //                    i = endtbl + 1;
        //                    title = bookmarks[i].Title;
        //                    skiplotbks = true;
        //                }
        //                if (InprogressFlag == "TOC" && isLOFExisted && isfigurebkmrksexist && !skiplofbks)
        //                {
        //                    i = endfig + 1;
        //                    title = bookmarks[i].Title;
        //                    skiplofbks = true;
        //                }
        //                title = bookmarks[i].Title;

        //                if (tocTitleFlag == false && isTOCCompleted == false && bookmarks[i].Level == 1 && (title.ToUpper() != "LIST OF TABLES") && title.ToUpper() != "LIST OF FIGURES" && rObj.isTOCExisted == false || (createtoc == true && tocTitleFlag == false && rObj.isTOCExisted == false && isTOCCompleted == false))
        //                {
        //                    InprogressFlag = "TOC";
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "TABLE OF CONTENTS";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;

        //                    tocPage = myDocumentNew.Pages.Insert(1);
        //                    //Set Page size for TOC pages
        //                    if (originalPG.Width > originalPG.Height)
        //                        myDocumentNew.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
        //                    else
        //                        myDocumentNew.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);

        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;

        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    tocTitleFlag = true;
        //                    if (createtoc == true && InprogressFlag == "TOC")
        //                    {
        //                        i = -1;
        //                    }
        //                }
        //                else if ((title.ToUpper() == "LIST OF TABLES"||(istablebkmrksexist && isLOTExisted)) && (isTOCCompleted || rObj.isTOCExisted == true) && lotFlag == false && rObj.isLOTExisted == false)
        //                {
        //                    InprogressFlag = "LOT";
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF TABLES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;

        //                    myDocumentNew.Save(sourcePath);
        //                    myDocumentNew.Dispose();
        //                    myDocumentNew = new Document(sourcePath);

        //                    if (tocTitleFlag == false && (rObj.isTOCExisted || isTOCCompleted))
        //                        tocTitleFlag = true;

        //                    if (LOTPagePosition != 0)
        //                    {
        //                        myDocumentNew.Pages.Insert(LOTPagePosition);
        //                        tocPage = myDocumentNew.Pages[LOTPagePosition];

        //                        //Set Page size for TOC pages
        //                        if (originalPG.Width > originalPG.Height)
        //                            myDocumentNew.Pages[LOTPagePosition].SetPageSize(originalPG.Height, originalPG.Width);
        //                        else
        //                            myDocumentNew.Pages[LOTPagePosition].SetPageSize(originalPG.Width, originalPG.Height);
        //                    }

        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;
        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    lotFlag = true;
        //                    if (istablebkmrksexist && isLOTExisted)
        //                        lotCount = lotbkcount;
        //                    else
        //                        lotCount = bookmarks[i].ChildItems.Count;
        //                    lofCount = 0;
        //                    flagCount = 0;
        //                }
        //                //else if (title.ToUpper() == "LIST OF FIGURES" && (((tocTitleFlag == true || rObj.isTOCExisted) && lotFlag == true && isLOTExisted) || ((tocTitleFlag == true||rObj.isTOCExisted) && lotFlag == false && !isLOTExisted)) && rObj.isLOFExisted == false)
        //                else if (((title.ToUpper() == "LIST OF FIGURES"||(isfigurebkmrksexist && isLOFExisted)) && ((lotFlag == true && isLOTExisted) || (lotFlag == false && !isLOTExisted) || (rObj.isTOCExisted && rObj.isLOTExisted && rObj.isLOFExisted == false)) && rObj.isLOFExisted == false) || (title.ToUpper() == "LIST OF FIGURES" || (isfigurebkmrksexist && isLOFExisted)) && rObj.isLOFExisted == false && isLOTExisted == false && isLOFExisted == true && (isTOCCompleted == true || rObj.isTOCExisted == true))
        //                {
        //                    InprogressFlag = "LOF";
        //                    titleFrag = new TextFragment();
        //                    titleFrag.Text = "LIST OF FIGURES";
        //                    titleFrag.TextState.LineSpacing = 20;
        //                    titleFrag.TextState.FontSize = 12;
        //                    titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
        //                    titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
        //                    titleFrag.TextState.FontStyle = FontStyles.Bold;

        //                    myDocumentNew.Save(sourcePath);
        //                    myDocumentNew.Dispose();
        //                    myDocumentNew = new Document(sourcePath);

        //                    //tocPage = myDocumentNew.Pages[LOTPgCnt + 1];
        //                    if (LOFPagePosition != 0)
        //                    {
        //                        myDocumentNew.Pages.Insert(LOFPagePosition);
        //                        tocPage = myDocumentNew.Pages[LOFPagePosition];

        //                        //Set Page size for TOC pages
        //                        if (originalPG.Width > originalPG.Height)
        //                            myDocumentNew.Pages[LOFPagePosition].SetPageSize(originalPG.Height, originalPG.Width);
        //                        else
        //                            myDocumentNew.Pages[LOFPagePosition].SetPageSize(originalPG.Width, originalPG.Height);
        //                    }
        //                    if (tocTitleFlag == false && (rObj.isTOCExisted || isTOCCompleted))
        //                        tocTitleFlag = true;

        //                    tocInfo = new TocInfo();
        //                    tocInfo.Title = titleFrag;
        //                    tocPage.TocInfo = tocInfo;

        //                    //tocPage.Paragraphs.Add(titleFrag);
        //                    if (isfigurebkmrksexist && isLOFExisted)
        //                        lofCount = lofbkcount;
        //                    else
        //                        lofCount = bookmarks[i].ChildItems.Count;
        //                    lofFlag = true;
        //                    flagCount = 0;
        //                    lotCount = 0;
        //                }


        //                if (tocTitleFlag == true && title.ToUpper() != "TABLE OF CONTENTS" && title.ToUpper() != "LIST OF TABLES" && title.ToUpper() != "LIST OF FIGURES" && i >= 0 && bookmarks[i].Level <= 4)
        //                {
        //                    if (lotCount > 0)
        //                    {
        //                        flagCount++;
        //                    }
        //                    if (lofCount > 0)
        //                    {
        //                        flagCount++;
        //                    }
        //                    Aspose.Pdf.Heading heading2 = new Heading(1);
        //                    heading2 = SetPropertiesForTOCItemsNew(rObj, bookmarks[i].Level, chLst);
        //                    //TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst);

        //                    TextSegment segment2 = new TextSegment();
        //                    heading2.TocPage = tocPage;
        //                    segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst, title);
        //                    heading2.TextState.ForegroundColor = Color.Blue;
        //                    heading2.TextState.LineSpacing = 10;
        //                    addpages = myDocumentNew.Pages.Count - initialPageCount;
        //                    // Destination page
        //                    heading2.DestinationPage = myDocumentNew.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber) + addpages]; //tocPages//myDocument.Pages[i + 2];                        
        //                    heading2.Top = myDocumentNew.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber) + addpages].Rect.Height;//tocPages
        //                    //LocalHyperlink lhl = new LocalHyperlink();
        //                    //lhl.TargetPageNumber = (bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber) + tocPages;
        //                    //segment2.Hyperlink = lhl;

        //                    if (bookmarks[i].Level == 2)
        //                    {
        //                        heading2.Margin.Left = 20;
        //                    }
        //                    else if (bookmarks[i].Level == 3)
        //                    {
        //                        heading2.Margin.Left = 25;
        //                    }
        //                    else if (bookmarks[i].Level == 4 || bookmarks[i].Level > 4)
        //                    {
        //                        heading2.Margin.Left = 30;
        //                    }

        //                    heading2.Segments.Add(segment2);
        //                    tocPage.Paragraphs.Add(heading2);

        //                }
        //                if (i == bookmarks.Count - 1 && InprogressFlag == "TOC")
        //                {
        //                    tocTitleFlag = true;
        //                    InprogressFlag = string.Empty;
        //                    isTOCCompleted = true;
        //                }
        //                //if (createtoc == true && InprogressFlag == "TOC" && isTOCCompleted == false)
        //                //{
        //                //    i = -1;
        //                //}
        //                if (lotCount > 0 && flagCount == lotCount && isLOFExisted == false)
        //                    break;
        //                else if (lofCount > 0 && flagCount == lofCount && flagCount > 0)
        //                    break;

        //                if (isLOTExisted == false && isLOFExisted == true && isTOCCompleted == true && tocTitleFlag == true && i == bookmarks.Count() - 1 && rObj.isLOFExisted == false)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                    tocTitleFlag = false;
        //                    InprogressFlag = string.Empty;
        //                }
        //                if (lotFlag == false && isLOTExisted == true && tocTitleFlag == true && i == bookmarks.Count() - 1 && rObj.isLOTExisted == false)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                    tocTitleFlag = false;
        //                    InprogressFlag = string.Empty;
        //                }
        //                else if (lotCount > 0 && flagCount == lotCount && rObj.isLOFExisted == false && isLOFExisted == true)
        //                {
        //                    i = -1;
        //                    flagCount = 0;
        //                    tocTitleFlag = false;
        //                    InprogressFlag = string.Empty;
        //                }
        //                else if (lofCount > 0 && flagCount == lofCount && flagCount > 0)
        //                {
        //                    break;
        //                }

        //            }
        //            myDocumentNew.Save(sourcePath);
        //            myDocumentNew.Dispose();

        //            bookmarkEditor = new PdfBookmarkEditor();
        //            //Open PDF file
        //            bookmarkEditor.BindPdf(sourcePath);
        //            bookmarks = bookmarkEditor.ExtractBookmarks();

        //            //Removing if already existed TOC bookmark
        //            for (int i = 0; i < bookmarks.Count(); i++)
        //            {
        //                if (bookmarks[i].Title.ToUpper() == "TABLE OF CONTENTS" && bookmarks[i].ChildItems.Count == 0)
        //                {
        //                    bookmarks.RemoveAt(i);
        //                    break;
        //                }
        //            }

        //            myDocumentNew = new Document(sourcePath);
        //            TextFragmentAbsorber textFragmentAbsorber = null;

        //            // Get the extracted text fragments
        //            bool setTOC = false;
        //            bool setLOT = false;
        //            bool setLOF = false;
        //            //Updating LOT and LOF links to new destination               
        //            for (int pn = 1; pn <= myDocumentNew.Pages.Count; pn++)
        //            {
        //                textFragmentAbsorber = new TextFragmentAbsorber();
        //                // Accept the absorber for all the pages
        //                myDocumentNew.Pages[pn].Accept(textFragmentAbsorber);

        //                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
        //                for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
        //                {
        //                    TextFragment tf = textFragmentCollection[frg];
        //                    if (tf.Text.ToUpper().Contains("LIST OF TABLES"))
        //                    {
        //                        Rectangle rect = tf.Rectangle;
        //                        Bookmark bookmarkLOT = new Bookmark();
        //                        //bookmarkTOC.Title = "TABLE OF CONTENTS";
        //                        for (int bk = 0; bk < bookmarks.Count; bk++)
        //                        {
        //                            if (bookmarks[bk].Level == 1 && bookmarks[bk].Title == "LIST OF TABLES")
        //                            {
        //                                bookmarkLOT = bookmarks[bk];
        //                                //bookmarkLOT.Level = 1;
        //                                bookmarkLOT.PageNumber = tf.Page.Number;
        //                                //bookmarkLOT.Action = "GoTo";
        //                                //bookmarkLOT.PageDisplay = "XYZ";
        //                                bookmarkLOT.PageDisplay_Left = (int)rect.LLX;
        //                                bookmarkLOT.PageDisplay_Top = (int)rect.URY;
        //                                //bookmarkLOT.PageDisplay_Zoom = 0;
        //                                bookmarks[bk] = bookmarkLOT;
        //                                setLOT = true;
        //                                break;
        //                            }
        //                        }
        //                        if (!setLOT && istablebkmrksexist && isLOTExisted)
        //                        {
        //                            bookmarkLOT.Title = "LIST OF TABLES";
        //                            bookmarkLOT.Level = 1;
        //                            bookmarkLOT.PageNumber = tf.Page.Number;
        //                            bookmarkLOT.Action = "GoTo";
        //                            bookmarkLOT.PageDisplay = "XYZ";
        //                            bookmarkLOT.PageDisplay_Left = (int)rect.LLX;
        //                            bookmarkLOT.PageDisplay_Top = (int)rect.URY;
        //                            bookmarkLOT.PageDisplay_Zoom = 0;
        //                            bookmarks.Insert(1, bookmarkLOT);
        //                            setLOT = true;
        //                        }
        //                    }
        //                    if (tf.Text.ToUpper().Contains("LIST OF FIGURES"))
        //                    {
        //                        Rectangle rect = tf.Rectangle;
        //                        Bookmark bookmarkLOF = new Bookmark();
        //                        //bookmarkTOC.Title = "TABLE OF CONTENTS";
        //                        for (int bk = 0; bk < bookmarks.Count; bk++)
        //                        {
        //                            if (bookmarks[bk].Level == 1 && bookmarks[bk].Title == "LIST OF FIGURES")
        //                            {
        //                                bookmarkLOF = bookmarks[bk];
        //                                //bookmarkLOF.Level = 1;
        //                                bookmarkLOF.PageNumber = tf.Page.Number;
        //                                //bookmarkLOF.Action = "GoTo";
        //                                //bookmarkLOF.PageDisplay = "XYZ";
        //                                bookmarkLOF.PageDisplay_Left = (int)rect.LLX;
        //                                bookmarkLOF.PageDisplay_Top = (int)rect.URY;
        //                                //bookmarkLOF.PageDisplay_Zoom = 0;
        //                                bookmarks[bk] = bookmarkLOF;
        //                                setLOF = true;
        //                                break;
        //                            }
        //                        }
        //                        if (!setLOF && isfigurebkmrksexist && isLOFExisted)
        //                        {
        //                            bookmarkLOF.Title = "LIST OF FIGURES";
        //                            bookmarkLOF.Level = 1;
        //                            bookmarkLOF.PageNumber = tf.Page.Number;
        //                            bookmarkLOF.Action = "GoTo";
        //                            bookmarkLOF.PageDisplay = "XYZ";
        //                            bookmarkLOF.PageDisplay_Left = (int)rect.LLX;
        //                            bookmarkLOF.PageDisplay_Top = (int)rect.URY;
        //                            bookmarkLOF.PageDisplay_Zoom = 0;
        //                            bookmarks.Insert(1, bookmarkLOF);
        //                            setLOF = true;
        //                        }
        //                    }
        //                    if (tf.Text.ToUpper().Contains("TABLE OF CONTENTS"))
        //                    {
        //                        Rectangle rect = tf.Rectangle;
        //                        Bookmark bookmarkTOC = new Bookmark();
        //                        bookmarkTOC.Title = "TABLE OF CONTENTS";
        //                        bookmarkTOC.Level = 1;
        //                        bookmarkTOC.PageNumber = tf.Page.Number;
        //                        bookmarkTOC.Action = "GoTo";
        //                        bookmarkTOC.PageDisplay = "XYZ";
        //                        bookmarkTOC.PageDisplay_Left = (int)rect.LLX;
        //                        bookmarkTOC.PageDisplay_Top = (int)rect.URY;
        //                        bookmarkTOC.PageDisplay_Zoom = 0;
        //                        bookmarks.Insert(0, bookmarkTOC);
        //                        setTOC = true;
        //                        //isTocLinkCreated = true;
        //                        //break;
        //                    }
        //                }
        //                if (setTOC && setLOT && setLOF)
        //                    break;
        //                //if (isTocLinkCreated)
        //                //    break;
        //            }

        //            bookmarkEditor.DeleteBookmarks();
        //            //Creating bookmarks
        //            for (int bk = 0; bk < bookmarks.Count; bk++)
        //            {
        //                if (bookmarks[bk].Level == 1)
        //                    bookmarkEditor.CreateBookmarks(bookmarks[bk]);
        //            }

        //            bookmarkEditor.Save(sourcePath);
        //            myDocumentNew.Dispose();

        //            ////myDocument = new Document(sourcePath);
        //            //PdfPageEditor pageEditor = new PdfPageEditor();
        //            //pageEditor.BindPdf(sourcePath);
        //            //PageSize originalPG = pageEditor.GetPageSize(tocPages + 1);
        //            //PageSize pz = null;
        //            //pageEditor.Close();
        //            //if (originalPG.Width > originalPG.Height)
        //            //{
        //            //    pz = new PageSize(originalPG.Height, originalPG.Width);
        //            //    originalPG = pz;
        //            //}

        //            //pageEditor = new PdfPageEditor();
        //            //pageEditor.BindPdf(sourcePath);
        //            //List<int> pgList = new List<int>();
        //            //for (int p = 1; p <= tocPages; p++)
        //            //{
        //            //    pgList.Add(p);
        //            //}
        //            //pageEditor.ProcessPages = pgList.ToArray();
        //            //pageEditor.PageSize = originalPG;
        //            //pageEditor.Save(sourcePath);
        //            //pageEditor.Close();

        //            //myDocumentNew = new Document(sourcePath);
        //            //int extraPagesCnt = myDocumentNew.Pages.Count;
        //            //for (int n = 0; n < (extraPagesCnt - (initialPageCount + tocPages)); n++)
        //            //{
        //            //    myDocumentNew.Pages.Delete(myDocumentNew.Pages.Count);
        //            //}
        //            ////if ((tocPages - PagesNeedToBeAdded) > 0)
        //            ////{
        //            ////    myDocumentNew.Pages.Delete((extraPagesCnt - initialPageCount));
        //            ////}
        //            //myDocumentNew.Save(sourcePath);
        //            //myDocumentNew.Dispose();

        //            rObj.Comments = "Table of contents created as per the bookmarks";
        //            rObj.QC_Result = "Fixed";
        //        }
        //        else
        //        {
        //            rObj.Comments = "No bookmarks existed in the document";
        //            rObj.QC_Result = "Passed";
        //        }
        //        rObj.CHECK_END_TIME = DateTime.Now;

        //    }
        //    catch (Exception ee)
        //    {
        //        rObj.Job_Status = "Error";
        //        rObj.QC_Result = "Error";
        //        rObj.Comments = "Technical error: " + ee.Message;
        //        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
        //    }
        //}


        public void CreateTOCFromBookmarksFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document document)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixToc = false;
            bool FixLot = false;
            bool FixLof = false;
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            List<string> bookmarkkName = new List<string>();
          
            //Open PDF file
            bookmarkEditor.BindPdf(document);
            //Extracting bookmarks from the source doc.
            Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
            //bookmarkEditor.Close();
            bookmarkEditor.Save(sourcePath);

            //Getting styles
            chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
            TextFragment headingStyle = new TextFragment();
            List<RegOpsQC> HeadingChecks = new List<RegOpsQC>();
            HeadingChecks = chLst.Where(x => x.Check_Name.Contains("Heading")).ToList();

            //Below logic to apply the heading style those are selected in the validation plan.        
            for (int fs = 0; fs < HeadingChecks.Count; fs++)
            {
                if (HeadingChecks[fs].Check_Name == "Heading Font Family" && chLst[fs].Check_Type == 1)
                {
                    headingStyle.TextState.Font = FontRepository.FindFont(chLst[fs].Check_Parameter);
                }
                if (HeadingChecks[fs].Check_Name == "Heading Font Style" && chLst[fs].Check_Type == 1)
                {
                    if (chLst[fs].Check_Parameter == "Bold")
                        headingStyle.TextState.FontStyle = FontStyles.Bold;
                    else if (chLst[fs].Check_Parameter == "Italic")
                        headingStyle.TextState.FontStyle = FontStyles.Italic;
                    else if (chLst[fs].Check_Parameter == "Regular")
                        headingStyle.TextState.FontStyle = FontStyles.Regular;
                }
                if (HeadingChecks[fs].Check_Name == "Heading Font Size" && chLst[fs].Check_Type == 1)
                    headingStyle.TextState.FontSize = float.Parse(chLst[fs].Check_Parameter);
            }
            
            try
            {
                //Checking for the TOC only check or fix
                bool fixornot = false;
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;

                    if (chLst[i].Check_Type == 1)
                    {
                        fixornot = true;
                    }
                }
                if (fixornot && bookmarks.Count > 0)
                {
                    //Loading source document
                    Document myDocument = new Document(sourcePath);
                    int initialPageCount = myDocument.Pages.Count();

                    //Getting page size of the original document.
                    PdfPageEditor pageEditor = new PdfPageEditor();
                    pageEditor.BindPdf(myDocument);
                    PageSize originalPG = pageEditor.GetPageSize(1);
                    //pageEditor.Close();

                    TocInfo tocInfo = new TocInfo();
                    Aspose.Pdf.Page tocPage = null;

                    TextFragment titleFrag = new TextFragment();
                    bool isLOTExisted = false;
                    bool isLOFExisted = false;
                    Guid guid = Guid.NewGuid();

                    bool IsLotBKExisted = false;
                    bool IsLofBKExisted = false;

                    bool isLOTPageIdentified = false;
                    bool isLOFPageIdentified = false;
                    bool isTOCPageIdentified = false;
                    int LOTPagePosition = 0;
                    int LOFPagePosition = 0;
                    string searchStr = string.Empty;
                    Regex regextbl = new Regex(@"(Table\s\d)", RegexOptions.IgnoreCase);
                    Regex regexfig = new Regex(@"(Figure\s\d)", RegexOptions.IgnoreCase);

                    Dictionary<int, Bookmark> dictLOT = new Dictionary<int, Bookmark>();
                    Dictionary<int, Bookmark> dictLOF = new Dictionary<int, Bookmark>();

                    //Seperating bookmarks that are related to TOC, LOT and LOF in different dictionary objects.               
                    for (int i = 0; i < bookmarks.Count(); i++)
                    {
                        if (bookmarks[i].Title.ToUpper() == "LIST OF TABLES" && bookmarks[i].ChildItems.Count > 0)
                        {
                            for (int k = i + 1; k <= i + bookmarks[i].ChildItems.Count; k++)
                            {
                                dictLOT.Add(k, bookmarks[k]);

                                if (k == i + bookmarks[i].ChildItems.Count)
                                    i = k;
                            }
                            IsLotBKExisted = true;
                            isLOTExisted = true;
                        }
                        else if (bookmarks[i].Title.ToUpper() == "LIST OF FIGURES" && bookmarks[i].ChildItems.Count > 0)
                        {
                            for (int k = i + 1; k <= i + bookmarks[i].ChildItems.Count; k++)
                            {
                                dictLOF.Add(k, bookmarks[k]);

                                if (k == i + bookmarks[i].ChildItems.Count)
                                    i = k;
                            }
                            IsLofBKExisted = true;
                            isLOFExisted = true;
                        }                        
                        else if (regextbl.IsMatch(bookmarks[i].Title))
                        {
                            dictLOT.Add(i, bookmarks[i]);
                            isLOTExisted = true;
                        }
                        else if (regexfig.IsMatch(bookmarks[i].Title))
                        {
                            dictLOF.Add(i, bookmarks[i]);
                            isLOFExisted = true;
                        }
                    }
                    int LotAndLofBKCount = 0;

                    //Below code to match the bookmark count
                    if (dictLOT.Keys.Count > 0)
                    {
                        LotAndLofBKCount = dictLOT.Keys.Count;
                        if (IsLotBKExisted)
                            LotAndLofBKCount = dictLOT.Keys.Count + 1;
                    }
                    if (dictLOF.Keys.Count > 0)
                    {
                        LotAndLofBKCount = LotAndLofBKCount + dictLOF.Keys.Count;
                        if (IsLofBKExisted)
                            LotAndLofBKCount = LotAndLofBKCount + 1;
                    }
                    //Below code is to check if TOC/LOT/LOF existed or not if not then will generate as per the standared order(i.e: TOC,LOT and LOF).
                    if (rObj.isTOCExisted == false && bookmarks.Count > LotAndLofBKCount)
                    {
                        searchStr = string.Empty;
                        bool isTOCTitleAdded = false;
                        /* Heading adding to the page */
                        TextFragment itleFrag = new TextFragment();
                        titleFrag.Text = "TABLE OF CONTENTS";
                        titleFrag.TextState.LineSpacing = 20;
                        titleFrag.TextState.FontSize = headingStyle.TextState.FontSize;
                        titleFrag.TextState.Font = headingStyle.TextState.Font;
                        titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                        titleFrag.TextState.FontStyle = headingStyle.TextState.FontStyle;

                        //Below code is checking where(in which page) to start TOC generation
                        if (rObj.isTOCExisted == false && rObj.isLOTExisted)
                            searchStr = "LIST OF TABLES";
                        else if (rObj.isTOCExisted == false && rObj.isLOTExisted == false && rObj.isLOFExisted)
                            searchStr = "LIST OF FIGURES";
                        else
                        {
                            isTOCPageIdentified = true;

                            tocPage = myDocument.Pages.Insert(1);
                            //Set Page size for TOC pages
                            if (originalPG.Width > originalPG.Height)
                                myDocument.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
                            else
                                myDocument.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);

                            //If LOT and LOF are not identified then considering the first page to create TOC.
                            tocInfo = new TocInfo();
                            tocInfo.IsShowPageNumbers = true;
                            tocInfo.Title = titleFrag;
                            tocPage.TocInfo = tocInfo;
                            isTOCTitleAdded = true;
                        }
                        if (isTOCPageIdentified == false)
                        {
                            //Need to verify where toc page ended
                            for (int k = 1; k <= myDocument.Pages.Count; k++)
                            {
                                using (MemoryStream textStream = new MemoryStream())
                                {                                   
                                    TextDevice textDevice = new TextDevice();
                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                    textDevice.ExtractionOptions = textExtOptions;
                                    textDevice.Process(myDocument.Pages[k], textStream);
                                    string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                    // Close memory stream
                                    textStream.Close();
                                    // Get text from memory stream                                    
                                    if (extractedText.ToUpper().Contains(searchStr))
                                    {
                                        tocPage = myDocument.Pages.Insert(k);

                                        //Set Page size for TOC pages
                                        if (originalPG.Width > originalPG.Height)
                                            myDocument.Pages[k].SetPageSize(originalPG.Height, originalPG.Width);
                                        else
                                            myDocument.Pages[k].SetPageSize(originalPG.Width, originalPG.Height);
                                        
                                        //Adding TOC title to the TOC page
                                        tocInfo = new TocInfo();
                                        tocInfo.IsShowPageNumbers = true;
                                        tocInfo.Title = titleFrag;
                                        tocPage.TocInfo = tocInfo;
                                        isTOCTitleAdded = true;
                                        isTOCPageIdentified = true;
                                    }
                                }
                                if (isTOCPageIdentified)
                                    break;
                            }
                        }
                        if (isTOCPageIdentified)
                        {
                            //If TOC page is identified then TOC generated.
                            for (int i = 0; i < bookmarks.Count(); i++)
                            {
                                if ((bookmarks[i].Title != "TABLE OF CONTENTS" && bookmarks[i].Title.ToUpper() != "LIST OF TABLES" && bookmarks[i].Title.ToUpper() != "LIST OF FIGURES" && bookmarks[i].Title.ToUpper() != "LIST OF TABLE" && bookmarks[i].Title.ToUpper() != "LIST OF FIGURE") && !dictLOT.ContainsKey(i) && !dictLOF.ContainsKey(i))
                                {
                                    if (isTOCTitleAdded == false)
                                    {
                                        tocPage = myDocument.Pages.Insert(1);

                                        //Set Page size for TOC pages
                                        if (originalPG.Width > originalPG.Height)
                                            myDocument.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
                                        else
                                            myDocument.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);

                                        //Adding TOC title to the TOC page
                                        tocInfo = new TocInfo();
                                        tocInfo.IsShowPageNumbers = true;
                                        tocInfo.Title = titleFrag;
                                        tocPage.TocInfo = tocInfo;
                                        isTOCTitleAdded = true;
                                    }
                                    /*---Below logic is to create TOC items in TOC page those are below level 5----*/
                                    if (bookmarks[i].Level <= 4)
                                    {                                        
                                        tocInfo.IsShowPageNumbers = true;
                                        Aspose.Pdf.Heading heading2 = new Heading(1);
                                        //Below method is to apply selected styles for the items in TOC
                                        heading2 = SetPropertiesForTOCItemsNew(rObj, bookmarks[i].Level, chLst);
                                        //TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst);

                                        TextSegment segment2 = new TextSegment();
                                        heading2.TocPage = tocPage;
                                        segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst, bookmarks[i].Title);
                                        heading2.TextState.ForegroundColor = Color.Blue;
                                        heading2.TextState.LineSpacing = 10;
                                        heading2.TextState.Font = segment2.TextState.Font;
                                        heading2.TextState.FontSize = segment2.TextState.FontSize;
                                        heading2.TextState.FontStyle = segment2.TextState.FontStyle;
                                        //insertedpgs = myDocument.Pages.Count - initialPageCount;
                                        // Setting link Destination
                                        heading2.DestinationPage = myDocument.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber + 1)]; //myDocument.Pages[i + 2];                        
                                        heading2.Top = myDocument.Pages[(bookmarks[i].PageNumber == 0 ? 1 : bookmarks[i].PageNumber + 1)].Rect.Height;                                        

                                        heading2.Segments.Add(segment2);
                                        tocPage.Paragraphs.Add(heading2);
                                        FixToc = true;
                                    }
                                }
                            }                          
                        }
                    }
                    // If LOT not existed then enters into below condition and finds the position where to start LOT page.
                    if (rObj.isLOTExisted == false)
                    {
                        searchStr = string.Empty;
                        if (dictLOT.Keys.Count > 0)
                        {
                            titleFrag = new TextFragment();
                            titleFrag.Text = "LIST OF TABLES";
                            titleFrag.TextState.LineSpacing = 20;
                            titleFrag.TextState.FontSize = headingStyle.TextState.FontSize;
                            titleFrag.TextState.Font = headingStyle.TextState.Font;
                            titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                            titleFrag.TextState.FontStyle = headingStyle.TextState.FontStyle;

                            if ((rObj.isTOCExisted == false && bookmarks.Count > LotAndLofBKCount) || rObj.isTOCExisted)
                            {
                                if (rObj.isTOCExisted)
                                    searchStr = "TABLE OF CONTENT";
                                else
                                {
                                    isLOTPageIdentified = true;
                                    tocInfo = new TocInfo();
                                    tocInfo.IsShowPageNumbers = true;
                                    tocInfo.Title = titleFrag;
                                    tocInfo.Title.IsInNewPage = true;
                                    tocPage.TocInfo.Title.IsInNewPage = true;
                                    tocPage.Paragraphs.Add(titleFrag);
                                }
                            }
                            else if (rObj.isLOTExisted == false && isLOTExisted && rObj.isLOFExisted)
                            {
                                if (rObj.isLOFExisted)
                                    searchStr = "LIST OF FIGURES";
                                else
                                {
                                    isLOTPageIdentified = true;
                                    tocInfo = new TocInfo();
                                    tocInfo.IsShowPageNumbers = true;
                                    tocInfo.Title = titleFrag;
                                    tocInfo.Title.IsInNewPage = true;
                                    tocPage.TocInfo.Title.IsInNewPage = true;
                                    tocPage.Paragraphs.Add(titleFrag);
                                }
                            }
                            else
                            {
                                LOTPagePosition = 1;
                                isLOTPageIdentified = true;

                                tocPage = myDocument.Pages.Insert(1);
                                //Set Page size for TOC pages
                                if (originalPG.Width > originalPG.Height)
                                    myDocument.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
                                else
                                    myDocument.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);

                                tocInfo = new TocInfo();
                                tocInfo.IsShowPageNumbers = true;
                                tocInfo.Title = titleFrag;
                                tocPage.TocInfo = tocInfo;
                            }
                            if (isLOTPageIdentified == false)
                            {
                                //Need to verify where toc page ended
                                for (int k = 1; k <= myDocument.Pages.Count; k++)
                                {
                                    using (MemoryStream textStream = new MemoryStream())
                                    {
                                        // Create text device
                                        TextDevice textDevice = new TextDevice();
                                        // Set text extraction options - set text extraction mode (Raw or Pure)
                                        Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                        textDevice.ExtractionOptions = textExtOptions;
                                        textDevice.Process(myDocument.Pages[k], textStream);
                                        string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                        // Close memory stream
                                        textStream.Close();
                                        // Get text from memory stream
                                        Regex regex = null;
                                        if (extractedText.ToUpper().Contains(searchStr))
                                        {
                                            for (int T = k; T <= myDocument.Pages.Count; T++)
                                            {
                                                using (MemoryStream textStreamTemp = new MemoryStream())
                                                {
                                                    // Create text device
                                                    textDevice = new TextDevice();
                                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                                    textExtOptions = new
                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                    textDevice.ExtractionOptions = textExtOptions;
                                                    textDevice.Process(myDocument.Pages[T], textStreamTemp);
                                                    extractedText = Encoding.Unicode.GetString(textStreamTemp.ToArray());
                                                    // Close memory stream  
                                                    textStreamTemp.Close();
                                                    regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                                    if ((extractedText.ToUpper().Contains("LIST OF TABLES") || extractedText.ToUpper().Contains("LIST OF FIGURES") || !regex.IsMatch(extractedText)) && !extractedText.ToUpper().Contains("TABLE OF CONTENT"))
                                                    {
                                                        LOTPagePosition = T;
                                                        //Inserting a new page for LOT after TOC.
                                                        tocPage = myDocument.Pages.Insert(T);
                                                        //Set Page size for TOC pages
                                                        if (originalPG.Width > originalPG.Height)
                                                            myDocument.Pages[T].SetPageSize(originalPG.Height, originalPG.Width);
                                                        else
                                                            myDocument.Pages[T].SetPageSize(originalPG.Width, originalPG.Height);

                                                        tocInfo = new TocInfo();
                                                        tocInfo.Title = titleFrag;
                                                        tocPage.TocInfo = tocInfo;

                                                        isLOTPageIdentified = true;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (isLOTPageIdentified)
                                        break;
                                }
                            }
                            //Once LOT page identified then enters into below code and generate LOT items from LOT items already stored in the dictionary obj.
                            if (isLOTPageIdentified)
                            {
                                foreach (var bk in dictLOT.Values)
                                {
                                    if (bk.Level <= 4)
                                    {
                                        Aspose.Pdf.Heading heading2 = new Heading(1);
                                        heading2 = SetPropertiesForTOCItemsNew(rObj, bk.Level, chLst);
                                        //TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst);

                                        TextSegment segment2 = new TextSegment();
                                        heading2.TocPage = tocPage;
                                        segment2 = SetPropertiesForTOCItems(rObj, bk.Level, chLst, bk.Title);
                                        heading2.TextState.ForegroundColor = Color.Blue;
                                        heading2.TextState.LineSpacing = 10;
                                        heading2.TextState.Font = segment2.TextState.Font;
                                        heading2.TextState.FontSize = segment2.TextState.FontSize;
                                        heading2.TextState.FontStyle = segment2.TextState.FontStyle;
                                        //insertedpgs = myDocument.Pages.Count - initialPageCount;
                                        // Destination page
                                        heading2.DestinationPage = myDocument.Pages[(bk.PageNumber == 0 ? 1 : bk.PageNumber + 1)]; //myDocument.Pages[i + 2];                        
                                        heading2.Top = myDocument.Pages[(bk.PageNumber == 0 ? 1 : bk.PageNumber + 1)].Rect.Height;
                                     
                                        heading2.Segments.Add(segment2);
                                        tocPage.Paragraphs.Add(heading2);
                                        FixLot = true;
                                    }

                                }                               
                            }
                        }
                    }
                    // If LOT not existed then enters into below condition and finds the position where to start LOT page.
                    if (rObj.isLOFExisted == false)
                    {
                        searchStr = string.Empty;
                        if (dictLOF.Keys.Count > 0)
                        {
                            titleFrag = new TextFragment();
                            titleFrag.Text = "LIST OF FIGURES";
                            titleFrag.TextState.LineSpacing = 20;
                            titleFrag.TextState.FontSize = headingStyle.TextState.FontSize;
                            titleFrag.TextState.Font = headingStyle.TextState.Font;
                            titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                            titleFrag.TextState.FontStyle = headingStyle.TextState.FontStyle;

                            if ((bookmarks.Count > LotAndLofBKCount && rObj.isTOCExisted==false) && rObj.isLOTExisted == false && isLOTExisted == false)
                            {                                
                                isLOFPageIdentified = true;
                                tocInfo = new TocInfo();
                                tocInfo.Title = titleFrag;
                                tocInfo.Title.IsInNewPage = true;
                                tocPage.TocInfo.Title.IsInNewPage = true;
                                tocPage.Paragraphs.Add(titleFrag);
                            }
                            else if (rObj.isLOTExisted || isLOTExisted)
                            {
                                if (rObj.isLOTExisted)
                                    searchStr = "LIST OF TABLES";
                                else
                                {
                                    isLOFPageIdentified = true;
                                    tocInfo = new TocInfo();
                                    tocInfo.Title = titleFrag;
                                    tocInfo.Title.IsInNewPage = true;
                                    tocPage.TocInfo.Title.IsInNewPage = true;
                                    tocPage.Paragraphs.Add(titleFrag);
                                }
                            }
                            else if(rObj.isTOCExisted)
                                searchStr = "TABLE OF CONTENT";                            
                            else
                            {
                                LOFPagePosition = 1;
                                isLOFPageIdentified = true;

                                tocPage = myDocument.Pages.Insert(1);
                                //Set Page size for TOC pages
                                if (originalPG.Width > originalPG.Height)
                                    myDocument.Pages[1].SetPageSize(originalPG.Height, originalPG.Width);
                                else
                                    myDocument.Pages[1].SetPageSize(originalPG.Width, originalPG.Height);

                                tocInfo = new TocInfo();
                                tocInfo.IsShowPageNumbers = true;
                                tocInfo.Title = titleFrag;
                                tocPage.TocInfo = tocInfo;
                            }
                            if (isLOFPageIdentified == false)
                            {
                                //Need to verify where LOT page ended
                                for (int k = 1; k <= myDocument.Pages.Count; k++)
                                {
                                    using (MemoryStream textStream = new MemoryStream())
                                    {
                                        // Create text device
                                        TextDevice textDevice = new TextDevice();
                                        // Set text extraction options - set text extraction mode (Raw or Pure)
                                        Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                        textDevice.ExtractionOptions = textExtOptions;
                                        textDevice.Process(myDocument.Pages[k], textStream);
                                        string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                        // Close memory stream
                                        textStream.Close();
                                        // Get text from memory stream
                                        Regex regex = null;
                                        if (extractedText.ToUpper().Contains(searchStr))
                                        {
                                            for (int T = k; T <= myDocument.Pages.Count; T++)
                                            {
                                                using (MemoryStream textStreamTemp = new MemoryStream())
                                                {
                                                    // Create text device
                                                    textDevice = new TextDevice();
                                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                                    textExtOptions = new
                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                    textDevice.ExtractionOptions = textExtOptions;
                                                    textDevice.Process(myDocument.Pages[T], textStreamTemp);
                                                    extractedText = Encoding.Unicode.GetString(textStreamTemp.ToArray());
                                                    // Close memory stream  
                                                    textStreamTemp.Close();
                                                    regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                                    if (extractedText.ToUpper().Contains("LIST OF FIGURES") || !regex.IsMatch(extractedText))
                                                    {
                                                        isLOFPageIdentified = true;
                                                        //Inserting a new page for LOF after TOC and LOT
                                                        //if(extractedText== "TABLE OF CONTENT")
                                                        tocPage = myDocument.Pages.Insert(T);
                                                        LOFPagePosition = T;

                                                        //Set Page size for TOC pages
                                                        if (originalPG.Width > originalPG.Height)
                                                            myDocument.Pages[T].SetPageSize(originalPG.Height, originalPG.Width);
                                                        else
                                                            myDocument.Pages[T].SetPageSize(originalPG.Width, originalPG.Height);

                                                        tocInfo = new TocInfo();
                                                        tocInfo.IsShowPageNumbers = true;
                                                        tocInfo.Title = titleFrag;
                                                        tocPage.TocInfo = tocInfo;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (isLOFPageIdentified)
                                        break;
                                }
                            }
                            //Once LOF page identified then enters into below code and generate LOT items from LOF items already stored in the dictionary obj.
                            if (isLOFPageIdentified)
                            {
                                foreach (var bk in dictLOF.Values)
                                {
                                    if (bk.Level <= 4)
                                    {
                                        Aspose.Pdf.Heading heading2 = new Heading(1);
                                        heading2 = SetPropertiesForTOCItemsNew(rObj, bk.Level, chLst);
                                        //TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level, chLst);

                                        TextSegment segment2 = new TextSegment();
                                        heading2.TocPage = tocPage;
                                        segment2 = SetPropertiesForTOCItems(rObj, bk.Level, chLst, bk.Title);
                                        heading2.TextState.ForegroundColor = Color.Blue;
                                        heading2.TextState.LineSpacing = 10;
                                        heading2.TextState.Font = segment2.TextState.Font;
                                        heading2.TextState.FontSize = segment2.TextState.FontSize;
                                        heading2.TextState.FontStyle = segment2.TextState.FontStyle;
                                        //insertedpgs = myDocument.Pages.Count - initialPageCount;
                                        // Destination page
                                        heading2.DestinationPage = myDocument.Pages[(bk.PageNumber == 0 ? 1 : bk.PageNumber + 1)]; //myDocument.Pages[i + 2];                        
                                        heading2.Top = myDocument.Pages[(bk.PageNumber == 0 ? 1 : bk.PageNumber + 1)].Rect.Height;
                                       
                                        heading2.Segments.Add(segment2);
                                        tocPage.Paragraphs.Add(heading2);
                                        FixLof = true;
                                        
                                    }
                                }                              
                            }
                        }
                    }
                    myDocument.Save(sourcePath);
                    //myDocument.Dispose();

                    //myDocument = new Document(sourcePath);
                    // New Code Ending

                    string InprogressFlag = string.Empty;
                    isLOTPageIdentified = false;
                    isLOFPageIdentified = false;
                    InprogressFlag = string.Empty;
                    //TextFragment titleFrag1 = new TextFragment();

                    guid = Guid.NewGuid();
                    //Loading the updated document for updating the bookmarks.                    
                    //Document myDocumentNew = new Document(sourcePath);
                    //tocInfo = new TocInfo();                    
                    //myDocumentNew.Save(sourcePath);
                    //myDocumentNew.Dispose();

                    bookmarkEditor = new PdfBookmarkEditor();
                    //Open PDF file
                    bookmarkEditor.BindPdf(sourcePath);
                    bookmarks = bookmarkEditor.ExtractBookmarks();

                    //Removing if already existed TOC bookmark
                    for (int i = 0; i < bookmarks.Count(); i++)
                    {
                        if (bookmarks[i].Title.ToUpper() == "TABLE OF CONTENTS" && bookmarks[i].ChildItems.Count == 0)
                        {
                            bookmarks.RemoveAt(i);
                            break;
                        }
                    }

                    Document myDocumentNew = new Document(sourcePath);
                    TextFragmentAbsorber textFragmentAbsorber = null;
                   
                    bool setTOC = false;
                    bool setLOT = false;
                    bool setLOF = false;

                    //Updating LOT and LOF bookmarks destination to new destination               
                    for (int pn = 1; pn <= myDocumentNew.Pages.Count; pn++)
                    {
                        textFragmentAbsorber = new TextFragmentAbsorber();                        
                        myDocumentNew.Pages[pn].Accept(textFragmentAbsorber);

                        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                        for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
                        {
                            TextFragment tf = textFragmentCollection[frg];
                            if (tf.Text.ToUpper().Contains("LIST OF TABLES")|| tf.Text.ToUpper().Contains("LIST OF TABLE"))
                            {
                                Rectangle rect = tf.Rectangle;
                                Bookmark bookmarkLOT = new Bookmark();
                                //bookmarkTOC.Title = "TABLE OF CONTENTS";
                                for (int bk = 0; bk < bookmarks.Count; bk++)
                                {
                                    if (bookmarks[bk].Title.ToUpper() == "LIST OF TABLES"|| bookmarks[bk].Title.ToUpper() == "LIST OF TABLE")
                                    {
                                        bookmarkLOT = bookmarks[bk];
                                        //bookmarkLOT.Level = 1;
                                        bookmarkLOT.PageNumber = tf.Page.Number;
                                        //bookmarkLOT.Action = "GoTo";
                                        //bookmarkLOT.PageDisplay = "XYZ";
                                        bookmarkLOT.PageDisplay_Left = (int)rect.LLX;
                                        bookmarkLOT.PageDisplay_Top = (int)rect.URY;
                                        //bookmarkLOT.PageDisplay_Zoom = 0;
                                        bookmarks[bk] = bookmarkLOT;
                                        setLOT = true;
                                        break;
                                    }
                                }
                                if (!setLOT && dictLOT.Count > 0 && isLOTExisted)
                                {
                                    bookmarkLOT.Title = "LIST OF TABLES";
                                    bookmarkLOT.Level = 1;
                                    bookmarkLOT.PageNumber = tf.Page.Number;
                                    bookmarkLOT.Action = "GoTo";
                                    bookmarkLOT.PageDisplay = "XYZ";
                                    bookmarkLOT.PageDisplay_Left = (int)rect.LLX;
                                    bookmarkLOT.PageDisplay_Top = (int)rect.URY;
                                    bookmarkLOT.PageDisplay_Zoom = 0;
                                    bookmarks.Insert(1, bookmarkLOT);
                                    setLOT = true;
                                }
                            }
                            if (tf.Text.ToUpper().Contains("LIST OF FIGURES")|| tf.Text.ToUpper().Contains("LIST OF FIGURE"))
                            {
                                Rectangle rect = tf.Rectangle;
                                Bookmark bookmarkLOF = new Bookmark();
                                //bookmarkTOC.Title = "TABLE OF CONTENTS";
                                for (int bk = 0; bk < bookmarks.Count; bk++)
                                {
                                    if (bookmarks[bk].Title.ToUpper() == "LIST OF FIGURES"|| bookmarks[bk].Title.ToUpper() == "LIST OF FIGURE")
                                    {
                                        bookmarkLOF = bookmarks[bk];
                                        //bookmarkLOF.Level = 1;
                                        bookmarkLOF.PageNumber = tf.Page.Number;
                                        //bookmarkLOF.Action = "GoTo";
                                        //bookmarkLOF.PageDisplay = "XYZ";
                                        bookmarkLOF.PageDisplay_Left = (int)rect.LLX;
                                        bookmarkLOF.PageDisplay_Top = (int)rect.URY;
                                        //bookmarkLOF.PageDisplay_Zoom = 0;
                                        bookmarks[bk] = bookmarkLOF;
                                        setLOF = true;
                                        break;
                                    }
                                }
                                if (!setLOF && dictLOF.Count > 0 && isLOFExisted)
                                {
                                    bookmarkLOF.Title = "LIST OF FIGURES";
                                    bookmarkLOF.Level = 1;
                                    bookmarkLOF.PageNumber = tf.Page.Number;
                                    bookmarkLOF.Action = "GoTo";
                                    bookmarkLOF.PageDisplay = "XYZ";
                                    bookmarkLOF.PageDisplay_Left = (int)rect.LLX;
                                    bookmarkLOF.PageDisplay_Top = (int)rect.URY;
                                    bookmarkLOF.PageDisplay_Zoom = 0;
                                    bookmarks.Insert(1, bookmarkLOF);
                                    setLOF = true;
                                }
                            }
                            if (tf.Text.ToUpper().Contains("TABLE OF CONTENTS") && setTOC == false)
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
                                setTOC = true;
                                //isTocLinkCreated = true;
                                //break;
                            }
                        }
                        if (setTOC && setLOT && setLOF)
                            break;
                        //if (isTocLinkCreated)
                        //    break;
                    }
                    //Deleting existed bookmarks
                    bookmarkEditor.DeleteBookmarks();
                    //Creating bookmarks as per updated document.
                    for (int bk = 0; bk < bookmarks.Count; bk++)
                    {
                        if (bookmarks[bk].Level == 1)
                            bookmarkEditor.CreateBookmarks(bookmarks[bk]);
                    }

                    bookmarkEditor.Save(sourcePath);
                    //myDocumentNew.Dispose();

                    //Correcting bookmarks order to TOC, LOT, LOF and Heading 1

                    bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(sourcePath);
                    //Extracting bookmarks from the source doc.
                    bookmarks = bookmarkEditor.ExtractBookmarks();
                    bookmarkEditor.DeleteBookmarks();
                    //Creating bookmarks                  
                    Bookmark bkNewTOC = new Bookmark();
                    Bookmark bkTempLOT = new Bookmark();
                    Bookmark bkTempLOF = new Bookmark();
                    Bookmarks bksNew = new Bookmarks();

                    int listoftablescount = 0;
                    int listoffigurescount = 0;

                
                    for (int bk = 0; bk < bookmarks.Count; bk++)
                    {
                        if (bookmarks[bk].Title.ToUpper() == "TABLE OF CONTENTS")
                        {
                            bkNewTOC = bookmarks[bk];
                            if(bookmarks[bk].ChildItems.Count>0)
                                bookmarks.RemoveRange(bk, bookmarks[bk].ChildItems.Count);                            
                            else
                                bookmarks.RemoveAt(bk);
                            bkNewTOC.Level = 1;
                            bk = bk - 1;
                        }
                        else if (bookmarks[bk].Title.ToUpper() == "LIST OF TABLES")
                        {
                            if (listoftablescount > 0)
                            {
                                if (bookmarks[bk].ChildItems.Count > 0)
                                {
                                    bkTempLOT.ChildItems.AddRange(bookmarks[bk].ChildItems);
                                }
                            }
                            else
                            {
                                bkTempLOT = bookmarks[bk];
                            }
                            
                            if (bookmarks[bk].ChildItems.Count > 0)
                                bookmarks.RemoveRange(bk, bookmarks[bk].ChildItems.Count);
                            else
                                bookmarks.RemoveAt(bk);
                            bkTempLOT.Level = 1;
                            bk = bk - 1;
                            listoftablescount++;
                            //bookmarks.Insert(1, bkTemp);
                        }
                        else if (bookmarks[bk].Title.ToUpper() == "LIST OF FIGURES")
                        {
                            if (listoffigurescount > 0)
                            {
                                if (bookmarks[bk].ChildItems.Count > 0)
                                {
                                    bkTempLOF.ChildItems.AddRange(bookmarks[bk].ChildItems);
                                }
                            }
                            else
                            {
                                bkTempLOF = bookmarks[bk];
                            }

                            if (bookmarks[bk].ChildItems.Count > 0)
                                bookmarks.RemoveRange(bk, bookmarks[bk].ChildItems.Count);
                            else
                                bookmarks.RemoveAt(bk);
                            bkTempLOF.Level = 1;
                            bk = bk - 1;
                            listoffigurescount++;
                            //bookmarks.Insert(2, bkTemp);
                        }                   
                    }
                    //Bookmarks bksNew = new Bookmarks();
                    if(bkNewTOC!=null)
                        bksNew.Insert(0, bkNewTOC);

                    if(bkTempLOT!=null && bkNewTOC != null)
                        bksNew.Insert(1, bkTempLOT);
                    else if (bkTempLOT != null && bkNewTOC == null)
                        bksNew.Insert(0, bkTempLOT);

                    if (bkTempLOF != null && bkTempLOT != null && bkNewTOC != null)
                        bksNew.Insert(2, bkTempLOF);
                    else if(bkTempLOF != null && (bkTempLOT == null || bkNewTOC == null))
                        bksNew.Insert(1, bkTempLOF);
                    else if (bkTempLOF != null && (bkTempLOT == null && bkNewTOC == null))
                        bksNew.Insert(0, bkTempLOF);

                    List<Bookmark> bookmarksTemp = bookmarks.Where(x => x.Level == 1).ToList();
                    for (int bk = 0; bk < bookmarksTemp.Count; bk++)
                    {
                        bksNew.Add(bookmarksTemp[bk]);
                    }
                    for (int bk = 0; bk < bksNew.Count; bk++)
                    {           
                        if(bksNew[bk].ChildItems.Count>0)
                        {                           
                            for (int ci = 0; ci < bksNew[bk].ChildItems.Count; ci++)
                            {
                                if (bksNew[bk].ChildItems[ci].Title.ToUpper() == "TABLE OF CONTENTS")
                                    bksNew[bk].ChildItems.Remove(bksNew[bk].ChildItems[ci]);
                                if (bksNew[bk].ChildItems[ci].Title.ToUpper() == "LIST OF TABLES")
                                    bksNew[bk].ChildItems.Remove(bksNew[bk].ChildItems[ci]);
                                if (bksNew[bk].ChildItems[ci].Title.ToUpper() == "LIST OF FIGURES")
                                    bksNew[bk].ChildItems.Remove(bksNew[bk].ChildItems[ci]);
                                if (bksNew[bk].ChildItems[ci].ChildItems.Count > 0)
                                {
                                    for (int di = 0; di < bksNew[bk].ChildItems[ci].ChildItems.Count; di++)
                                    {
                                        if (bksNew[bk].ChildItems[ci].ChildItems[di].Title.ToUpper() == "TABLE OF CONTENTS")
                                            bksNew[bk].ChildItems[ci].ChildItems.Remove(bksNew[bk].ChildItems[ci].ChildItems[di]);
                                        if (bksNew[bk].ChildItems[ci].ChildItems[di].Title.ToUpper() == "LIST OF TABLES")
                                            bksNew[bk].ChildItems[ci].ChildItems.Remove(bksNew[bk].ChildItems[ci].ChildItems[di]);
                                        if (bksNew[bk].ChildItems[ci].ChildItems[di].Title.ToUpper() == "LIST OF FIGURES")
                                            bksNew[bk].ChildItems[ci].ChildItems.Remove(bksNew[bk].ChildItems[ci].ChildItems[di]);

                                    }

                                }
                            }                            
                        }             
                        if (bksNew[bk].Level == 1)
                            bookmarkEditor.CreateBookmarks(bksNew[bk]);
                    }
                    //Saving final document
                    bookmarkEditor.Save(sourcePath);
                    //myDocumentNew.Dispose();
                    //bookmarkEditor.Close();
                    document = new Document(sourcePath);
                    //document.Save(sourcePath);
                
                    if (rObj.Comments == "TOC,LOT exist but LOF does not exist")
                    {
                        if(!FixToc && !FixLot && FixLof)
                        {
                            rObj.Comments = "TOC,LOT exist but LOF does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else
                            rObj.Comments = "TOC,LOT exist but LOF does not exist";
                    }
                    else if (rObj.Comments == "TOC exist but LOT and LOF does not exist")
                    {
                        if(!FixToc && FixLot && FixLof)
                        {
                            rObj.Comments = "TOC exist but LOT and LOF does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if(!FixToc && FixLot && !FixLof)
                        {
                            rObj.Comments = "TOC exist but LOT does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if(!FixToc && !FixLot && FixLof)
                        {
                            rObj.Comments = "TOC exist but LOF does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if(FixToc && FixLot && FixLof)
                        {
                            rObj.Comments = "TOC, LOT and LOF do not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else
                        rObj.Comments = "TOC exist but LOT and LOF does not exist";                      
                    }
                    else if (rObj.Comments == "LOT and LOF exist but not TOC")
                    {
                        if (FixToc && !FixLot && !FixLof)
                        {
                            rObj.Comments = "LOT and LOF exist but not TOC. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else
                            rObj.Comments = "LOT and LOF exist but not TOC";                      
                    }
                    else if (rObj.Comments == "LOF exists but not TOC and LOT")
                    {
                        if (FixToc && FixLot && !FixLof)
                        {
                            rObj.Comments = "LOF exists but not TOC and LOT. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if (FixToc && !FixLot && !FixLof)
                        {
                            rObj.Comments = "LOF exists but not TOC. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if (!FixToc && FixLot && !FixLof)
                        {
                            rObj.Comments = "LOF exists but not LOT. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else
                            rObj.Comments = "LOF exists but not TOC and LOT";                    
                    }
                    else if (rObj.Comments == "TOC, LOT and LOF do not exist")
                    {
                        if(FixToc && FixLot && FixLof)
                        {
                            rObj.Comments = "TOC, LOT and LOF do not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else
                        rObj.Comments = "TOC, LOT and LOF do not exist";
                    }
                    else if (rObj.Comments == "TOC, LOT do not exist")
                    {
                        if (FixToc && FixLot && !FixLof)
                        {
                            rObj.Comments = "TOC, LOT do not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if(!FixToc && FixLot && !FixLof)
                        {
                            rObj.Comments = "LOT does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if (FixToc && !FixLot && !FixLof)
                        {
                            rObj.Comments = "TOC does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else
                            rObj.Comments = "TOC, LOT do not exist";
                    }
                    else if (rObj.Comments == "TOC, LOF do not exist")
                    {
                        if (FixToc && !FixLot && FixLof)
                        {
                            rObj.Comments = "TOC, LOF do not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if (FixToc && !FixLot && !FixLof)
                        {
                            rObj.Comments = "TOC does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else if (!FixToc && !FixLot && FixLof)
                        {
                            rObj.Comments = "LOF does not exist. Fixed";
                            rObj.Is_Fixed = 1;
                        }
                        else
                            rObj.Comments = "TOC, LOF do not exist";                      
                    }
                    else if(rObj.Comments == "TOC does not exist" && FixToc)
                    {
                        rObj.Comments = "TOC does not exist. Fixed";
                        rObj.Is_Fixed = 1;
                    }
                }               
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
        /// Embeded File 
        /// </summary>
        /// <param name="rObj"></param>

        public void EmbeddedFilecheck(RegOpsQC rObj, string path,Document document)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);
                bool audiovideo = false;
                bool javascriptkeys = false;
                IList keys = (System.Collections.IList)document.JavaScript.Keys;
                if (keys.Count > 0)
                    javascriptkeys = true;
                for (int i = 1; i <= document.Pages.Count; i++)
                {
                    var mediaAnnotations = document.Pages[i].Annotations
                   .Where(a => (a.AnnotationType == AnnotationType.Screen)
                   || (a.AnnotationType == AnnotationType.Sound)
                   || (a.AnnotationType == AnnotationType.RichMedia))
                   .Cast<Annotation>();
                    foreach (var ma in mediaAnnotations)
                    {
                        audiovideo = true;
                    }
                }
                if (document.EmbeddedFiles.Count > 0)
                {
                    rObj.Comments = "File embedded in document";
                    rObj.QC_Result = "Failed";
                }
                if (javascriptkeys == true)
                {
                    if (document.EmbeddedFiles != null)
                    {
                        rObj.Comments = rObj.Comments + " and java script keys present in document";
                        rObj.QC_Result = "Failed";
                    }
                    else
                    {
                        rObj.Comments = "java script keys present in document";
                        rObj.QC_Result = "Failed";
                    }
                }
                if (audiovideo == true)
                {
                    if (document.EmbeddedFiles != null || javascriptkeys == true)
                    {
                        rObj.Comments = rObj.Comments + " and Audio video clips are present in document";
                        rObj.QC_Result = "Failed";
                    }
                    else
                    {
                        rObj.Comments = "Audio video clips are present in document";
                        rObj.QC_Result = "Failed";
                    }

                }
                if (document.EmbeddedFiles.Count == 0  && javascriptkeys == false && audiovideo == false)
                {
                    //rObj.Comments = "No File Embeded and no audio video clips in document";
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
        /// Embeded File 
        /// </summary>
        /// <param name="rObj"></param>
        public void FixEmbeddedFilecheck(RegOpsQC rObj, string path, Document document)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);
                bool flag = false;
                IList keys = (System.Collections.IList)document.JavaScript.Keys;
                for (int i = 0; i <= keys.Count; i++)
                {
                    keys.Remove(i);
                    flag = true;
                }
                for (int i = 1; i <= document.Pages.Count; i++)
                {
                    var mediaAnnotations = document.Pages[i].Annotations
                   .Where(a => (a.AnnotationType == AnnotationType.Screen)
                   || (a.AnnotationType == AnnotationType.Sound)
                   || (a.AnnotationType == AnnotationType.RichMedia))
                   .Cast<Annotation>();
                    foreach (var ma in mediaAnnotations)
                    {
                        document.Pages[i].Annotations.Delete(ma);
                        flag = true;
                    }
                }
                if (document.EmbeddedFiles != null || flag == true)
                {
                    if (document.EmbeddedFiles != null)
                        document.EmbeddedFiles.Delete();
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.Comments = "No File Embeded in document";
                    rObj.QC_Result = "Passed";
                }
               // document.Save(sourcePath);
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
        ///Hyperlink includes all blue text-check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void HyperlinkIncludesAllBlueText(RegOpsQC rObj,  Document doc)
        {
            try
            {
                
                rObj.CHECK_START_TIME = DateTime.Now;
                //Document doc = new Document(sourcePath);
                bool isFailed = false;
                List<List<TextFragment>> bluetexts = BluetextLinkcheck(doc);
                string pageNumbers = "";
                if (bluetexts.Count > 0)
                {
                    foreach (List<TextFragment> lst in bluetexts)
                    {
                        if (lst.Count == 1)
                        {
                            Page page = lst[0].Page;
                            AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                            page.Accept(selector);
                            IList<Annotation> list = selector.Selected;
                            if (list.Count > 0)
                            {
                                foreach (LinkAnnotation annot in list)
                                {
                                    TextFragmentAbsorber lta = new TextFragmentAbsorber();
                                    Rectangle rect = annot.Rect;
                                    lta.TextSearchOptions = new TextSearchOptions(annot.Rect);
                                    lta.Visit(page);
                                    string content = "";
                                    foreach (TextFragment tf in lta.TextFragments)
                                    {
                                        content = content + tf.Text;
                                    }
                                    if (lst[0].Text.Contains(content) && lst[0].Text != content)
                                    {
                                        if (pageNumbers == "")
                                        {
                                            pageNumbers = page.Number.ToString() + ", ";
                                        }
                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                        isFailed = true;
                                    }


                                }

                            }
                        }
                        else if (lst.Count > 1)
                        {
                            string combinedtext = "";
                            foreach (TextFragment tf in lst)
                                combinedtext = combinedtext + tf.Text;
                            Rectangle recttext = new Rectangle(lst[0].Rectangle.LLX, lst[0].Rectangle.LLY, lst[lst.Count - 1].Rectangle.URX, lst[lst.Count - 1].Rectangle.URY);
                            Page page = lst[0].Page;
                            AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                            page.Accept(selector);
                            IList<Annotation> list = selector.Selected;
                            if (list.Count > 0)
                            {
                                foreach (LinkAnnotation annot in list)
                                {
                                    TextFragmentAbsorber lta = new TextFragmentAbsorber();
                                    Rectangle rect = annot.Rect;
                                    lta.TextSearchOptions = new TextSearchOptions(annot.Rect);
                                    lta.Visit(page);
                                    string content = "";
                                    foreach (TextFragment tf in lta.TextFragments)
                                    {
                                        content = content + tf.Text;
                                    }
                                    if (combinedtext.Contains(content) && combinedtext != content)
                                    {
                                        if (pageNumbers == "")
                                        {

                                            pageNumbers = page.Number.ToString() + ", ";
                                        }
                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";

                                        isFailed = true;
                                    }

                                }

                            }
                        }
                    }
                }
                if (isFailed)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Hyperlink(s) do not include all blue text in: " + pageNumbers.Trim().TrimEnd(',');
                    rObj.CommentsWOPageNum = "Hyperlink(s) do not include all blue text";
                    rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Hyperlinks includes all blue text.";
                }
                //doc.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);

            }
        }
        /// <summary>
        /// Hyperlink Includes All BlueText Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void HyperlinkIncludesAllBlueTextFix(RegOpsQC rObj,  Document doc)
        {
            try
            {
               // Document doc = new Document(sourcePath);
                bool Fixed = false;
                rObj.FIX_START_TIME = DateTime.Now;
                List<List<TextFragment>> bluetexts = BluetextLinkcheck(doc);
                if (bluetexts.Count > 0)
                {
                    foreach (List<TextFragment> lst in bluetexts)
                    {
                        if (lst.Count == 1)
                        {
                            Page page = lst[0].Page;
                            AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                            page.Accept(selector);
                            IList<Annotation> list = selector.Selected;
                            if (list.Count > 0)
                            {
                                foreach (LinkAnnotation annot in list)
                                {
                                    TextFragmentAbsorber lta = new TextFragmentAbsorber();
                                    Rectangle rect = annot.Rect;
                                    lta.TextSearchOptions = new TextSearchOptions(annot.Rect);
                                    lta.Visit(page);
                                    string content = "";
                                    foreach (TextFragment tf in lta.TextFragments)
                                    {
                                        content = content + tf.Text;
                                    }
                                    if (lst[0].Text.Contains(content))
                                    {
                                        var rct = lst[0].Rectangle;
                                        annot.Rect = rct;
                                        var link = annot;
                                        page.Annotations.Delete(annot);
                                        LinkAnnotation linkannot = new LinkAnnotation(page, rct);
                                        linkannot = link;
                                        page.Annotations.Add(linkannot);
                                        Fixed = true;
                                    }
                                }

                            }
                        }
                        else if (lst.Count > 1)
                        {
                            string combinedtext = "";
                            foreach (TextFragment tf in lst)
                                combinedtext = combinedtext + tf.Text;
                            Rectangle recttext = new Rectangle(lst[0].Rectangle.LLX, lst[0].Rectangle.LLY, lst[lst.Count - 1].Rectangle.URX, lst[lst.Count - 1].Rectangle.URY);
                            Page page = lst[0].Page;
                            AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                            page.Accept(selector);
                            IList<Annotation> list = selector.Selected;
                            if (list.Count > 0)
                            {
                                foreach (LinkAnnotation annot in list)
                                {
                                    TextFragmentAbsorber lta = new TextFragmentAbsorber();
                                    Rectangle rect = annot.Rect;
                                    lta.TextSearchOptions = new TextSearchOptions(annot.Rect);
                                    lta.Visit(page);
                                    string content = "";
                                    foreach (TextFragment tf in lta.TextFragments)
                                    {
                                        content = content + tf.Text;
                                    }
                                    if (combinedtext.Contains(content))
                                    {
                                        annot.Rect = recttext;
                                        var link = annot;
                                        page.Annotations.Delete(annot);
                                        LinkAnnotation linkannot = new LinkAnnotation(page, recttext);
                                        linkannot = link;
                                        page.Annotations.Add(linkannot);
                                        Fixed = true;
                                    }
                                }

                            }
                        }
                    }
                }
                if (Fixed)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Hyperlinks includes all blue text.";
                }
                //doc.Save(sourcePath);
                //doc.Dispose();
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);

            }
        }

        public List<List<TextFragment>> BluetextLinkcheck(Document doc)
        {
            List<List<TextFragment>> txtfragarr = new List<List<TextFragment>>();
            List<TextFragment> bluetexts = new List<TextFragment>();
            foreach (Page page in doc.Pages)
            {
                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                page.Accept(ta);
                TextFragmentCollection TextFrgmtColl = ta.TextFragments;


                for (int tf = 1; tf <= TextFrgmtColl.Count; tf++)
                {
                    int tfFrtcount = 0;
                    // double hposition = 0;
                    List<TextFragment> tflst = new List<TextFragment>();

                    if (TextFrgmtColl[tf].TextState.ForegroundColor == Color.Blue)
                    {
                        tfFrtcount = tf;
                        tflst.Add(TextFrgmtColl[tf]);
                        int count = 0;
                        if (tf + 1 < TextFrgmtColl.Count)
                        {
                            for (int tfnxt = tf + 1; tfnxt <= TextFrgmtColl.Count; tfnxt++)
                            {

                                if (TextFrgmtColl[tfnxt].TextState.ForegroundColor == Color.Blue)
                                {
                                    count++;
                                    tflst.Add(TextFrgmtColl[tfnxt]);
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
                    if (tflst.Count > 0)
                        txtfragarr.Add(tflst);


                }
            }
            return txtfragarr;

        }
        public void RedactReplaceNextCheck(RegOpsQC rObj, List<RegOpsQC> chLst, Document doc)
        {
            try
            {
                string textinput = "";
                string[] textvalues = { };
                int pageNo =0;
                List<int> commentpage = new List<int>();
                List<string> Flag = new List<string>();
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;

                    if (chLst[i].Check_Name.ToString() == "Text")
                    {
                        textinput = chLst[i].Check_Parameter.ToString();
                        textvalues = textinput.Split(',');
                    }

                    if (chLst[i].Check_Name.ToString() == "Page No" && chLst[i].Check_Parameter.ToString()!="")
                    {
                        pageNo = Convert.ToInt32(chLst[i].Check_Parameter.ToString());
                    }

                }
                if (textvalues.Length != 0)
                {
                    foreach (string values in textvalues)
                    {
                        string name = values;
                        int page = pageNo;
                        //Regex re = new Regex(@"(?=((name\s).?\w.*?\s))", RegexOptions.IgnoreCase);
                        Regex re1 = new Regex(name + "\\s\\w.*?\\s", RegexOptions.IgnoreCase);
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(re1);
                        if (pageNo != 0)
                        {
                            foreach (Page p in doc.Pages)
                            {
                                if(p.Number == page)
                                {
                                    p.Accept(textFragmentAbsorber);
                                    //textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(p.Rect);
                                    foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                    {
                                        //Console.WriteLine(pp.Text);
                                        Rectangle rec = pp.Rectangle;
                                        string temp = pp.Text;
                                        temp = temp.Trim();
                                        string[] dd = temp.Split(' ');
                                        string final = dd[1];
                                        TextFragmentAbsorber text1 = new TextFragmentAbsorber(final);
                                        text1.TextSearchOptions = new TextSearchOptions(rec);
                                        p.Accept(text1);
                                        if (text1.TextFragments.Count > 0)
                                        {
                                            Flag.Add("passed");
                                            commentpage.Add(page);
                                        }
                                        else
                                        {
                                            Flag.Add("Failed");
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (Page p in doc.Pages)
                            {
                                p.Accept(textFragmentAbsorber);
                                //textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(p.Rect);
                                foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                {
                                    //Console.WriteLine(pp.Text);
                                    Rectangle rec = pp.Rectangle;
                                    string temp = pp.Text;
                                    temp = temp.Trim();
                                    string[] dd = temp.Split(' ');
                                    string final = dd[1];
                                    TextFragmentAbsorber text1 = new TextFragmentAbsorber(final);
                                    text1.TextSearchOptions = new TextSearchOptions(rec);
                                    p.Accept(text1);
                                    if (text1.TextFragments.Count > 0)
                                    {
                                        Flag.Add("passed");
                                        commentpage.Add(p.Number);
                                    }
                                    else
                                    {
                                        Flag.Add("Failed");
                                    }
                                }
                            }
                        }
                    }
                    List<int> lst2 = commentpage.Distinct().ToList();
                    if (Flag.Contains("passed"))
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Redact string failed at :" + Pagenumber;
                        rObj.CommentsWOPageNum = "Redact present in the Document";
                        rObj.PageNumbersLst = lst2;
                    }
                    else
                    {
                        rObj.QC_Result = "passed";
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
        }
        public void RedactReplaceNextFix(RegOpsQC rObj, List<RegOpsQC> chLst, Document doc)
        {
            try
            {
                string textinput = "";
                string[] textvalues = { };
                int pageNo = 0;
                bool Flag = false;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;

                    if (chLst[i].Check_Name.ToString() == "Text")
                    {
                        textinput = chLst[i].Check_Parameter.ToString();
                        textvalues = textinput.Split(',');
                    }

                    if (chLst[i].Check_Name.ToString() == "Page No")
                    {
                        pageNo = Convert.ToInt32(chLst[i].Check_Parameter.ToString());
                    }

                }
                if (textvalues.Length != 0)
                {
                    foreach (string values in textvalues)
                    {
                        string name = values;
                        int page = pageNo;
                        //Regex re = new Regex(@"(?=((name\s).?\w.*?\s))", RegexOptions.IgnoreCase);
                        Regex re1 = new Regex(name + "\\s\\w.*?\\s", RegexOptions.IgnoreCase);
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(re1);
                        if (pageNo != 0)
                        {
                            foreach (Page p in doc.Pages)
                            {
                                if (p.Number == page)
                                {
                                    p.Accept(textFragmentAbsorber);
                                    //textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(p.Rect);
                                    foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                    {
                                        //Console.WriteLine(pp.Text);
                                        Rectangle rec = pp.Rectangle;
                                        string temp = pp.Text;
                                        temp = temp.Trim();
                                        string[] dd = temp.Split(' ');
                                        string final = dd[1];
                                        TextFragmentAbsorber text1 = new TextFragmentAbsorber(final);
                                        text1.TextSearchOptions = new TextSearchOptions(rec);
                                        p.Accept(text1);
                                        foreach (TextFragment a in text1.TextFragments)
                                        {
                                            Rectangle finalrec = a.Rectangle;
                                            RedactionAnnotation annot = new RedactionAnnotation(p, finalrec);
                                            annot.FillColor = Aspose.Pdf.Color.Black;
                                            annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                                            annot.Repeat = true;
                                            doc.Pages[p.Number].Annotations.Add(annot);
                                            annot.Redact();
                                            Flag = true;
                                        }

                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (Page p in doc.Pages)
                            {
                                p.Accept(textFragmentAbsorber);
                                //textFragmentAbsorber.TextSearchOptions = new TextSearchOptions(p.Rect);
                                foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                {
                                    //Console.WriteLine(pp.Text);
                                    Rectangle rec = pp.Rectangle;
                                    string temp = pp.Text;
                                    temp = temp.Trim();
                                    string[] dd = temp.Split(' ');
                                    string final = dd[1];
                                    TextFragmentAbsorber text1 = new TextFragmentAbsorber(final);
                                    text1.TextSearchOptions = new TextSearchOptions(rec);
                                    p.Accept(text1);
                                    foreach (TextFragment a in text1.TextFragments)
                                    {
                                        Rectangle finalrec = a.Rectangle;
                                        RedactionAnnotation annot = new RedactionAnnotation(p, finalrec);
                                        annot.FillColor = Aspose.Pdf.Color.Black;
                                        annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                                        annot.Repeat = true;
                                        doc.Pages[p.Number].Annotations.Add(annot);
                                        annot.Redact();
                                        Flag = true;
                                    }

                                }
                            }
                        }
                        
                    }
                    if (Flag == true)
                    {
                        rObj.Is_Fixed = 1;
                        rObj.Comments = rObj.Comments + ". Fixed";
                    }
                }
            }
            catch(Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }

            
        }

        public void RedactGivenTextCheck(RegOpsQC rObj, List<RegOpsQC> chLst, Document doc)
        {
            try
            {
                string textinput = "";
                string[] textvalues = { };
                int pageNo = 0;
                List<int> commentpage = new List<int>();
                List<string> Flag = new List<string>();
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;

                    if (chLst[i].Check_Name.ToString() == "Text")
                    {
                        textinput = chLst[i].Check_Parameter.ToString();
                        textvalues = textinput.Split(',');
                    }

                    if (chLst[i].Check_Name.ToString() == "Page No" && chLst[i].Check_Parameter.ToString() != "")
                    {
                        pageNo = Convert.ToInt32(chLst[i].Check_Parameter.ToString());
                    }

                }
                if (textvalues.Length != 0)
                {
                    foreach (string values in textvalues)
                    {
                        string name = values;
                        int page = pageNo;
                        //Regex re1 = new Regex("\\s" + name + "\\s", RegexOptions.IgnoreCase);
                        Regex re1 = new Regex("(\\s+"+name+"\\s|^"+name+ "\\s|\\s"+name+"$)", RegexOptions.IgnoreCase);
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(re1);
                        if (pageNo != 0)
                        {
                            foreach (Page p in doc.Pages)
                            {
                                if (p.Number == page)
                                {
                                    p.Accept(textFragmentAbsorber);
                                    foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                    {
                                        Rectangle rec = pp.Rectangle;
                                        string temp = pp.Text;
                                        temp = temp.Trim();
                                        TextFragmentAbsorber text = new TextFragmentAbsorber(temp);
                                        text.TextSearchOptions = new TextSearchOptions(rec);
                                        p.Accept(text);
                                        if (text.TextFragments.Count > 0)
                                        {
                                            Flag.Add("passed");
                                            commentpage.Add(page);
                                        }
                                        else
                                        {
                                            Flag.Add("Failed");
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (Page p in doc.Pages)
                            {
                                p.Accept(textFragmentAbsorber);
                                foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                {
                                    Rectangle rec = pp.Rectangle;
                                    string temp = pp.Text;
                                    temp = temp.Trim();
                                    TextFragmentAbsorber text = new TextFragmentAbsorber(temp);
                                    text.TextSearchOptions = new TextSearchOptions(rec);
                                    p.Accept(text);
                                    if (text.TextFragments.Count > 0)
                                    {
                                        Flag.Add("passed");
                                        commentpage.Add(p.Number);
                                    }
                                    else
                                    {
                                        Flag.Add("Failed");
                                    }
                                }
                            }
                        }
                    }
                    List<int> lst2 = commentpage.Distinct().ToList();
                    if (Flag.Contains("passed"))
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Redact string failed at :" + Pagenumber;
                        rObj.CommentsWOPageNum = "Redact present in the Document";
                        rObj.PageNumbersLst = lst2;
                    }
                    else
                    {
                        rObj.QC_Result = "passed";
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
        }
        public void RedactGivenTextFix(RegOpsQC rObj, List<RegOpsQC> chLst, Document doc)
        {
            try
            {
                string textinput = "";
                string[] textvalues = { };
                int pageNo = 0;
                bool Flag = false;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;

                    if (chLst[i].Check_Name.ToString() == "Text")
                    {
                        textinput = chLst[i].Check_Parameter.ToString();
                        textvalues = textinput.Split(',');
                    }

                    if (chLst[i].Check_Name.ToString() == "Page No")
                    {
                        pageNo = Convert.ToInt32(chLst[i].Check_Parameter.ToString());
                    }

                }
                if (textvalues.Length != 0)
                {
                    foreach (string values in textvalues)
                    {
                        string name = values;
                        int page = pageNo;
                        //Regex re1 = new Regex("\\s" + name + "\\s", RegexOptions.IgnoreCase);
                        Regex re1 = new Regex("(\\s+" + name + "\\s|^" + name + "\\s|\\s" + name + "$)", RegexOptions.IgnoreCase);

                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(re1);
                        if (pageNo != 0)
                        {
                            foreach (Page p in doc.Pages)
                            {
                                if (p.Number == page)
                                {
                                    p.Accept(textFragmentAbsorber);
                                    foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                    {
                                        Rectangle rec = pp.Rectangle;
                                        string temp = pp.Text;
                                        temp = temp.Trim();
                                        TextFragmentAbsorber text = new TextFragmentAbsorber(temp);
                                        text.TextSearchOptions = new TextSearchOptions(rec);
                                        p.Accept(text);
                                        foreach (TextFragment a in text.TextFragments)
                                        {
                                            Rectangle finalrec = a.Rectangle;
                                            RedactionAnnotation annot = new RedactionAnnotation(p, finalrec);
                                            annot.FillColor = Aspose.Pdf.Color.Black;
                                            annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                                            annot.Repeat = true;
                                            doc.Pages[p.Number].Annotations.Add(annot);
                                            annot.Redact();
                                            Flag = true;
                                        }

                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (Page p in doc.Pages)
                            {
                                p.Accept(textFragmentAbsorber);
                                foreach (TextFragment pp in textFragmentAbsorber.TextFragments)
                                {
                                    Rectangle rec = pp.Rectangle;
                                    string temp = pp.Text;
                                    temp = temp.Trim();
                                    TextFragmentAbsorber text = new TextFragmentAbsorber(temp);
                                    text.TextSearchOptions = new TextSearchOptions(rec);
                                    p.Accept(text);
                                    foreach (TextFragment a in text.TextFragments)
                                    {
                                        Rectangle finalrec = a.Rectangle;
                                        RedactionAnnotation annot = new RedactionAnnotation(p, finalrec);
                                        annot.FillColor = Aspose.Pdf.Color.Black;
                                        annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                                        annot.Repeat = true;
                                        doc.Pages[p.Number].Annotations.Add(annot);
                                        annot.Redact();
                                        Flag = true;
                                    }

                                }
                            }
                        }

                    }
                    if (Flag == true)
                    {
                        rObj.Is_Fixed = 1;
                        rObj.Comments = rObj.Comments + ". Fixed";
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


        }

    }
}