using iText.IO.Font;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Colorspace;
using iText.Layout;
using iText.Layout.Element;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFCreator
{
    public class PDFGenerator
    {
        Random rnd = new Random();
        string path = "";
        PdfWriter writer;
        PdfDocument pdf;
        Document document;
        float pageWidth;
        float pageHeight;
        Color bbsBlue;
        
        public PDFGenerator()
        {
            PageSize ps = PageSize.A4;
            pageWidth = ps.GetWidth();
            pageHeight = ps.GetHeight();
            bbsBlue = new DeviceRgb(0, 97, 160);
            GenerateFilePath();
        }

        public void SpoolPDF()
        {
            CreateNewPDF();
        }


        private void CreateNewPDF()
        {
            writer = new PdfWriter(path);
            pdf = new PdfDocument(writer);
            document = new Document(pdf);

            SpoolPageOne();               

            document.Close();
        }

        private void SpoolPageOne()
        {            
            //title            
            string text = GetText(MergeField.Title);
            Paragraph title = new Paragraph(text);
            float titleParaHeight = 126;
            float titleParaWidth = 200;
            title.SetFontSize(13);
            title.SetFontColor(bbsBlue);
            title.SetWidth(titleParaWidth);
            title.SetHeight(titleParaHeight);
            title.SetFixedPosition((pageWidth / 2 - titleParaWidth), pageHeight - titleParaHeight, titleParaWidth);
            document.Add(title);
        }

        #region helper methods
        private void AddParagraph(Document doc, Paragraph paragraph)
        {
            doc.Add(paragraph);
        }

        private void GenerateFilePath()
        {
            path = @"\\bbs-actuaries\dfsdata\users\hodsonl\pdfs\test";
            path += rnd.Next(1, 1000);
            path += ".pdf";
        }

        private void ModifyExistingDocument()
        {
            string oldFile = @"\\bbs-actuaries\dfsdata\users\hodsonl\pdfs\Template.pdf";
            string newFile = path;

            // open the reader       
            PdfDocument pdfDocument = new PdfDocument(new PdfReader(oldFile), new PdfWriter(newFile));
            Document document = new Document(pdfDocument);

            Paragraph para = new Paragraph("This is a test!!!!!!!!");
            PageSize ps = PageSize.A4;
            var width = ps.GetWidth();
            var height = ps.GetHeight();
            para.SetRelativePosition(200, 200, 300, 200);
            AddParagraph(document, para);
            document.Close();
        }

        private string GetText(MergeField field)
        {
            var text = "";
            switch (field)
            {
                case MergeField.Title:
                    text = "The Baptist Pension Scheme (BPS) – «EMPLOYER_NAME»";
                    break;
                case MergeField.SubTitle:
                    Console.WriteLine("Case 2");
                    break;
                default:
                    //oops
                    break;                
            }
            return text;
        }

        private enum MergeField
        {
            Title,
            SubTitle
        }
        #endregion
    }
}
