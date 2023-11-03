using CMCai.Models;
using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Data;

namespace CMCai.Actions
{
    public class PDFGenerator
    {

        protected void AddPageNumber(string folderPath)
        {

            try
            {
                string currentYear = DateTime.Now.Year.ToString();
                byte[] bytes = File.ReadAllBytes(folderPath);
                Font blackFont = FontFactory.GetFont("Times New Roman", 7, 1, BaseColor.BLACK);
                string filename = string.Empty;
                using (MemoryStream stream = new MemoryStream())
                {
                    PdfReader reader = new PdfReader(bytes);
                    using (PdfStamper stamper = new PdfStamper(reader, stream))
                    {
                        int pages = reader.NumberOfPages;

                        DateTime PrintTime = DateTime.Now;
                        TimeZone zone = TimeZone.CurrentTimeZone;

                        string Time = PrintTime + "(IST)";
                        for (int i = 1; i <= pages; i++)
                        {
                           // ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new Phrase("Testing", blackFont), 25f, 14f, 0);
                            //ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_LEFT, new Phrase(" Generated on:" + Time + "  Copyright © " + currentYear + " DDi. All Rights Reserved. " + "Page  " + (i) + " of " + pages, blackFont), 25f, 14f, 0);
                            ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_LEFT, new Phrase("                                                                                                                                                                                                                                                                 Page | " + (i), blackFont), 25f, 14f, 0);
                        }
                    }
                    bytes = stream.ToArray();
                }
                File.WriteAllBytes(folderPath, bytes);

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }


        }
        public string RegistrationReportPDF(string htmlpdfsrtng, string CRval)
        {
            string filename = string.Empty;
            string folderPath = CRval + ".pdf";
            try
            {
                // string folderPath = AppDomain.CurrentDomain.BaseDirectory + "ReportDocs\\" + CRval + ".pdf";

                FileStream fs = new FileStream(folderPath, FileMode.Create);
                Document document = new Document(PageSize.A4, 40, 40, 30, 30);
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                writer.PageEvent = new itextEvents();
                document.Open();
                HTMLWorker hw = new HTMLWorker(document);
                string fontpath = AppDomain.CurrentDomain.BaseDirectory + "/Content1/";
                string arialuniTff = Path.Combine(fontpath, "Times_New_Romance.ttf");
                iTextSharp.text.FontFactory.Register(arialuniTff);
                iTextSharp.text.html.simpleparser.StyleSheet ST = new iTextSharp.text.html.simpleparser.StyleSheet();
                ST.LoadTagStyle(HtmlTags.BODY, HtmlTags.FACE, "Times New Roman Unicode MS");
                ST.LoadTagStyle(HtmlTags.BODY, HtmlTags.ENCODING, BaseFont.IDENTITY_H);
                ST.LoadTagStyle(HtmlTags.BODY, HtmlTags.STYLE, "font-size:10px;text-align:left; font-family:Times New Roman;");
                ST.LoadTagStyle(HtmlTags.P, HtmlTags.STYLE, "font-size:8px;text-align:left;font-weight: lighter;font-family:Times New Roman;");
                ST.LoadTagStyle(HtmlTags.SPAN, HtmlTags.STYLE, "font-size:10px;text-align:left;font-family:Times New Roman;");
                ST.LoadTagStyle(HtmlTags.TD, HtmlTags.BORDER, "0.5");
                ST.LoadTagStyle(HtmlTags.TD, HtmlTags.STYLE, "font-size:7px;text-align:center;");
                ST.LoadTagStyle(HtmlTags.TH, HtmlTags.BORDER, "0.5");
                ST.LoadTagStyle(HtmlTags.TH, HtmlTags.BGCOLOR, "#b2beb5");
                ST.LoadTagStyle(HtmlTags.TH, HtmlTags.STYLE, "font-size:8px;text-align:center;");
                
                //ST.LoadTagStyle(HtmlTags.TABLE, HtmlTags.BGCOLOR, "#ff");
                ST.LoadTagStyle(HtmlTags.TABLE, HtmlTags.STYLE, "font-size:8px;text-align:center;height:20px;width:50px;");
                List<IElement> list = HTMLWorker.ParseToList(new StringReader(htmlpdfsrtng), ST);
                foreach (var element in list)
                {
                    document.Add(element);
                }
                document.Close();
                writer.Close();
                fs.Close();
                return folderPath;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "FAILED";
            }
            finally
            {
                AddPageNumber(folderPath);
            }

        }

        public class itextEvents : IPdfPageEvent
        {
            //Create object of PdfContentByte
            PdfContentByte pdfContent;
            public void OnEndPage(iTextSharp.text.pdf.PdfWriter writer, iTextSharp.text.Document document)
            {
                //We are going to add two strings in header. Create separate Phrase object with font setting and string to be included
                Phrase p1Header = new Phrase("Generated on" + System.DateTime.Now + " (" + TimeZone.CurrentTimeZone + ")", FontFactory.GetFont("Times New Roman", 10));
                Phrase p2Header = new Phrase("Copyright © 2017 DDi. All Rights Reserved.", FontFactory.GetFont("Times New Roman", 10));
                //create iTextSharp.text Image object using local image path
                //iTextSharp.text.Image imgPDFRight = iTextSharp.text.Image.GetInstance(AppDomain.CurrentDomain.BaseDirectory + "/external/Images/DDi-small.png");
                //iTextSharp.text.Image imgPDFLeft = iTextSharp.text.Image.GetInstance(AppDomain.CurrentDomain.BaseDirectory + "/external/Images/visu-logo.png");
                //imgPDFRight.ScalePercent(22f);
                //imgPDFRight.SetAbsolutePosition(150f, 180f);
                //imgPDFLeft.ScalePercent(32f);
                //imgPDFRight.SetAbsolutePosition(150f, 180f);
                ////Create PdfTable object
                //PdfPTable pdfTab = new PdfPTable(2);
                ////We will have to create separate cells to include image logo and 2 separate strings
                //PdfPCell pdfCell1 = new PdfPCell(imgPDFLeft);
                //PdfPCell pdfCell3 = new PdfPCell(imgPDFRight);
                //set the alignment of all three cells and set border to 0
                //pdfCell1.HorizontalAlignment = Element.ALIGN_CENTER;
                //pdfCell1.PaddingRight = 180;
                //pdfCell3.HorizontalAlignment = Element.ALIGN_RIGHT;
                //pdfCell1.Border = 0;
                //pdfCell3.Border = 0;
                ////add all three cells into PdfTable
                //pdfTab.AddCell(pdfCell1);
                //pdfTab.AddCell(pdfCell3);
                //pdfTab.TotalWidth = document.PageSize.Width - 48;
                ////call WriteSelectedRows of PdfTable. This writes rows from PdfWriter in PdfTable
                ////first param is start row. -1 indicates there is no end row and all the rows to be included to write
                ////Third and fourth param is x and y position to start writing
                //pdfTab.WriteSelectedRows(0, -1, 10, document.PageSize.Height - 15, writer.DirectContent);
                ////set pdfContent value
                pdfContent = writer.DirectContent;
                //Move the pointer and draw line to separate header section from rest of page
                pdfContent.MoveTo(30, document.PageSize.Height - 35);
                //pdfContent.LineTo(document.PageSize.Width - 40, document.PageSize.Height - 35);
                pdfContent.Stroke();
            }


            void IPdfPageEvent.OnChapter(PdfWriter writer, Document document, float paragraphPosition, Paragraph title)
            {
            }
            void IPdfPageEvent.OnChapterEnd(PdfWriter writer, Document document, float paragraphPosition)
            {
            }
            void IPdfPageEvent.OnCloseDocument(PdfWriter writer, Document document)
            {
            }

            void IPdfPageEvent.OnGenericTag(PdfWriter writer, Document document, Rectangle rect, string text)
            {
            }
            void IPdfPageEvent.OnOpenDocument(PdfWriter writer, Document document)
            {
            }
            void IPdfPageEvent.OnParagraph(PdfWriter writer, Document document, float paragraphPosition)
            {
            }
            void IPdfPageEvent.OnParagraphEnd(PdfWriter writer, Document document, float paragraphPosition)
            {
            }
            void IPdfPageEvent.OnSection(PdfWriter writer, Document document, float paragraphPosition, int depth, Paragraph title)
            {
            }
            void IPdfPageEvent.OnSectionEnd(PdfWriter writer, Document document, float paragraphPosition)
            {
            }
            void IPdfPageEvent.OnStartPage(PdfWriter writer, Document document)
            {
            }
          
        }
       
    }
    public class _events : PdfPageEventHelper
    {
        public void pdfPrint(DataTable dt)
         {
            string filename = string.Empty;
            try
            {
                //Document Doc;
                filename = "Audit Trail" + "_on_" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy@HH.mm.ss") + ".pdf";
                HttpContext.Current.Response.ContentType = "application/pdf";
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=\"" + filename + "\"");
                HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                 Document Doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 130f, 20f);
                 Doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 130f, 20f);
                MemoryStream ms = new MemoryStream();
                 _events HeaderEvent = new _events();
                PdfWriter pw = PdfWriter.GetInstance(Doc, HttpContext.Current.Response.OutputStream);
                 pw.PageEvent = HeaderEvent;
                Doc.Open();
                Font fnt = FontFactory.GetFont("Arial, Helvetica, sans-serif", 10);
                PdfPTable PdfTable = new PdfPTable(dt.Columns.Count);
                PdfPCell PdfPCell = null;
                PdfTable.WidthPercentage = 100;
                string str = string.Empty;
                for (int rows = 0; rows < dt.Rows.Count; rows++)
                {
                    if (rows == 0)
                    {
                        for (int column = 0; column < dt.Columns.Count; column++)
                        {
                            PdfPCell = new PdfPCell(new Phrase(new Chunk(dt.Columns[column].ColumnName.ToString(), FontFactory.GetFont("Arial", 10, Font.BOLD,BaseColor.BLACK))));
                            PdfPCell.HorizontalAlignment = 1;
                            PdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                            PdfPCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                            PdfTable.AddCell(PdfPCell);
                        }
                    }
                    for (int column = 0; column < dt.Columns.Count; column++)
                    {
                        if (!string.IsNullOrEmpty(dt.Rows[rows][column].ToString()))
                            PdfPCell = new PdfPCell(new Phrase(new Chunk(dt.Rows[rows][column].ToString(), fnt)));
                        else
                            PdfPCell = new PdfPCell(new Phrase(new Chunk("abcd", fnt)));
                        PdfTable.AddCell(PdfPCell);
                        PdfTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        PdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    }
                }
                PdfTable.Summary = "Test data";
                PdfTable.HeaderRows = 1;
                Doc.Add(PdfTable);
                Doc.Close();
                HttpContext.Current.Response.Write(Doc);
                HttpContext.Current.Response.End();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}