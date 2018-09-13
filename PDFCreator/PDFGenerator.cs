using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
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
        
        public PDFGenerator()
        {
            PageSize ps = PageSize.A4;
            pageWidth = ps.GetWidth();
            pageHeight = ps.GetHeight();
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
            string title = "The Baptist Pension Scheme (BPS) – «EMPLOYER_NAME»";
            Paragraph titlePara = new Paragraph(title);
            float titleParaHeight = 126;
            float titleParaWidth = 144;
            titlePara.SetFontSize(13);
            titlePara.SetFont("Arial");
            titlePara.SetWidth(titleParaWidth);
            titlePara.SetHeight(titleParaHeight);
            titlePara.SetFixedPosition((pageWidth / 2 - titleParaWidth), pageHeight - titleParaHeight, titleParaWidth);
            document.Add(titlePara);
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
        #endregion
    }
}
