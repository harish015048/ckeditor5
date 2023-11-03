using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Bytescout.PDF2HTML;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Web;

namespace CMCai.Actions
{
    public class DocumentComparisonUtil
    {
        public string m_PreviewDocumentObj = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"]+ "\\DCM\\";
        /// <summary>
        /// Compare the two documents using Aspose.Words and save the result as a Word document
        /// </summary>
        /// <param name="document1">First document</param>
        /// <param name="document2">Second document</param>
        /// <param name="comparisonDocument">Comparison document</param>
        public void Compare(string document1, string document2, string comparisonDocument, ref int added, ref int deleted)
        {
            added = 0;
            deleted = 0;

            // Load both documents in Aspose.Words
            Aspose.Words.Document doc1 = new Aspose.Words.Document(document1);
            Aspose.Words.Document doc2 = new Aspose.Words.Document(document2);
            Aspose.Words.Comparing.CompareOptions compareOptions = new Aspose.Words.Comparing.CompareOptions();
            compareOptions.IgnoreFields = false;
            compareOptions.IgnoreHeadersAndFooters = false;
            compareOptions.IgnoreFormatting = false;            
            if (doc1.Revisions.Count == 0 && doc2.Revisions.Count == 0)
                doc1.Compare(doc2, "a", DateTime.Now,compareOptions);

            foreach (Revision revision in doc1.Revisions)
            {
                switch (revision.RevisionType)
                {
                    case RevisionType.Insertion:
                        added++;
                        break;
                    case RevisionType.Deletion:
                        deleted++;
                        break;
                }
            }
            DocumentBuilder builder = new DocumentBuilder(doc1);            
            Debug.WriteLine("Revisions: " + doc1.Revisions.Count);          
            doc1.Save(comparisonDocument);
            builder.InsertField(Aspose.Words.Fields.FieldType.FieldPage, false);
            HtmlFixedSaveOptions option = new HtmlFixedSaveOptions();
            option.ExportEmbeddedImages = true;
            option.ExportEmbeddedFonts = true;
            option.ExportEmbeddedCss = true;
            option.ShowPageBorder = true;
            option.UpdateFields = false;         
            doc1.Save(comparisonDocument.Trim() + "html/Aspose_DocToHTML.html", option);
        }

        public void PreviewDoctoHtml(string previewDocument, ref int added, ref int deleted)
        {
            added = 0;
            deleted = 0;
            string comppathexe = Path.GetExtension(previewDocument).Remove(0, 1);
            string prefilename = Path.GetFileNameWithoutExtension(previewDocument);
            if (comppathexe != "pdf")
            {
                Document doc1 = new Document(previewDocument);
                doc1.Save(previewDocument);
                HtmlSaveOptions option = new HtmlSaveOptions(SaveFormat.Html);
                option.ExportImagesAsBase64 = true;
                doc1.Save(previewDocument.Trim() + "html/Aspose_DocToHTML.html", option);
            }
            else
            {
               
                // Create Bytescout.PDF2HTML.HTMLExtractor instance
                HTMLExtractor extractor = new HTMLExtractor();
                extractor.RegistrationName = "demo";
                extractor.RegistrationKey = "demo";

                // Set HTML with CSS extraction mode
                extractor.ExtractionMode = HTMLExtractionMode.HTMLWithCSS;
                
                // Load sample PDF document
                extractor.LoadDocumentFromFile(m_PreviewDocumentObj + prefilename + ".pdf");

                extractor.ExtractAnnotations = true;
                // Set HTML with CSS extraction mode
                extractor.ExtractionMode = HTMLExtractionMode.HTMLWithCSS;
                // Embed images into HTML file
                extractor.SaveImages = ImageHandling.Embed;
                // Save extracted HTML to file
                extractor.SaveHtmlToFile(m_PreviewDocumentObj + prefilename + "_Bytescout_PdfToHTML.html");
               
                extractor.Dispose();
            }
        }           
        }
}