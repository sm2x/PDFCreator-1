using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Colorspace;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
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
        Random _rnd = new Random();
        string _path = "";
        PdfWriter _writer;
        PdfDocument _pdf;
        Document _document;
        float _pageWidth;
        float _pageHeight;
        Color _baptistBlue;
        float _paraOffset;
        float _horizontalOffset;
        DataAccess _dataAccess;


        public PDFGenerator()
        {
            PageSize ps = PageSize.A4;
            _pageWidth = ps.GetWidth();
            _pageHeight = ps.GetHeight();
            _paraOffset = -30;
            _horizontalOffset = 130;
            _baptistBlue = new DeviceRgb(0, 176, 240);
            _dataAccess = new DataAccess(true);
            GenerateFilePath();
        }

        public void SpoolPDF()
        {
            CreateNewPDF();
        }


        private void CreateNewPDF()
        {
            _writer = new PdfWriter(_path);
            _pdf = new PdfDocument(_writer);
            _document = new Document(_pdf);
            _document.SetMargins(0, 100, 0, 100);

            SpoolPageOne();               

            _document.Close();
        }

        private void SpoolPageOne()
        {
            float titleParaHeight;
            float titleParaWidth;
            string text = "";
            float verticalPosition = _pageHeight;

            //logo            
            string strFilePath = @"\\bbs-actuaries\dfsdata\users\hodsonl\Visual Studio 2017\Projects\PDFCreator\PDFCreator\Images\baptist.png";
            ImageData l = ImageDataFactory.Create(strFilePath);
            Image img = new Image(l);
            img.SetHeight(39);
            img.SetWidth(110);
            img.SetFixedPosition(_pageWidth - 180, _pageHeight - 65);
            _document.Add(img);

            //title            
            text = GetText(MergeField.Title);
            Paragraph title = new Paragraph(text);
            titleParaHeight = 126;
            titleParaWidth = _pageWidth;
            title.SetFontSize(13);
            title.SetBold();
            title.SetTextAlignment(TextAlignment.CENTER);
            title.SetFontColor(_baptistBlue);
            title.SetUnderline();
            title.SetWidth(titleParaWidth);
            title.SetHeight(titleParaHeight);
            verticalPosition -= 220;
            title.SetFixedPosition((_pageWidth - titleParaWidth), verticalPosition, titleParaWidth);
            _document.Add(title);            

            //estimated employer debt
            text = GetText(MergeField.EstimatedEmployerDebt);
            Paragraph employerDebt = new Paragraph(text);
            titleParaHeight = 126;
            titleParaWidth = _pageWidth;
            employerDebt.SetFontSize(13);
            employerDebt.SetUnderline();
            employerDebt.SetBold();
            employerDebt.SetTextAlignment(TextAlignment.CENTER);
            employerDebt.SetFontColor(_baptistBlue);
            employerDebt.SetWidth(titleParaWidth);
            employerDebt.SetHeight(titleParaHeight);
            verticalPosition -= 30;
            employerDebt.SetFixedPosition((_pageWidth - titleParaWidth), verticalPosition, titleParaWidth);
            _document.Add(employerDebt);

            //"Introduction"
            text = "Introduction";
            Paragraph intro = new Paragraph(text);
            titleParaHeight = 126;
            titleParaWidth = _pageWidth;
            intro.SetFontSize(11);
            intro.SetTextAlignment(TextAlignment.CENTER);
            intro.SetFontColor(_baptistBlue);
            intro.SetUnderline();
            intro.SetBold();
            intro.SetWidth(titleParaWidth);
            intro.SetHeight(titleParaHeight);
            verticalPosition -= 40;
            intro.SetFixedPosition((_pageWidth - titleParaWidth), verticalPosition, titleParaWidth);
            _document.Add(intro);

            //List of items            
            List introList = new List(ListNumberingType.DECIMAL);                                    
            titleParaWidth = _pageWidth - _horizontalOffset;
            introList.SetFontSize(10);
            introList.SetTextAlignment(TextAlignment.LEFT);
            introList.SetWidth(titleParaWidth);
            introList.SetMinHeight(150);
            verticalPosition -= 300;

            //Item 1
            Paragraph para = new Paragraph();
            para.Add(new Text(GetText(MergeField.IntroListItemOnePartOne)).SetBold());
            para.Add(new Text(GetText(MergeField.IntroListItemOnePartTwo)));

            ListItem listItemOne = new ListItem();
            listItemOne.Add(para);
            listItemOne.SetPaddingBottom(10);            
            introList.Add(listItemOne);
            
            //Item 2            
            ListItem listItemTwo = new ListItem(GetText(MergeField.IntroListItemTwo));            
            listItemTwo.SetPaddingBottom(10);
            introList.Add(listItemTwo);           

            //Item 3  
            ListItem listItemThree = new ListItem(GetText(MergeField.IntroListItemThree));
            listItemThree.SetPaddingBottom(10);
            introList.Add(listItemThree);

            //Item 4
            ListItem listItemFour = new ListItem(GetText(MergeField.IntroListItemFour));
            listItemFour.SetPaddingBottom(10);
            introList.Add(listItemFour);

            //Item 5
            ListItem listItemFive = new ListItem(GetText(MergeField.IntroListItemFive));
            listItemFive.SetPaddingBottom(10);
            introList.Add(listItemFive);

            //item 5 nested list
            List item5SubList = new List(ListNumberingType.ZAPF_DINGBATS_2);
            item5SubList.Add("an employer debt only becomes due when an employer incurs a “cessation event”");
            item5SubList.Add("a cessation event normally only occurs when an employer stops employing any active members of the BPS");            
            item5SubList.SetFontSize(10);
            item5SubList.SetFixedPosition((_pageWidth - titleParaWidth - 50), verticalPosition + 115, titleParaWidth);
            _document.Add(item5SubList);
            

            //Item 6
            ListItem listItemSix = new ListItem(GetText(MergeField.IntroListItemSix));
            listItemSix.SetPaddingBottom(10);
            listItemSix.SetPaddingTop(40);
            introList.Add(listItemSix);

            //item 7
            ListItem listItemSeven = new ListItem(GetText(MergeField.IntroListItemSeven));
            listItemSeven.SetPaddingBottom(10);
            introList.Add(listItemSeven);                     

            introList.SetFixedPosition((_pageWidth - titleParaWidth - (_horizontalOffset / 2)), verticalPosition, titleParaWidth);
            _document.Add(introList);

            //nested list
            List item7SubList = new List(ListNumberingType.ZAPF_DINGBATS_2);
            item7SubList.Add(string.Format("The total due would be greater (potentially significantly greater) if your organisation has previously incurred a separate cessation event. ", _dataAccess.GetPreviousCEWord()));
            item7SubList.Add("The amount due could be higher or lower if your organisation is currently in a “period of grace” (a “period of grace” applies if you have had a cessation event and you do not currently employ an active member, but you have confirmed that you expect to take on a new employee who will become a BPS member soon).  As a period of grace can only be granted if an employer requests it, your organisation should be aware if a period of grace applies in your case.  If an employer in a period of grace does not take on a new active member before the end of a period of grace, that employer’s debt will be calculated based on the finances of the Scheme at the date of the cessation event rather than at a current date.");
            verticalPosition -= 140;
            item7SubList.SetFontSize(10);
            item7SubList.SetFixedPosition((_pageWidth - titleParaWidth - 50), verticalPosition, titleParaWidth);
            _document.Add(item7SubList);


        }

        #region helper methods
        private void AddParagraph(Document doc, Paragraph paragraph)
        {
            doc.Add(paragraph);
        }

        private void GenerateFilePath()
        {
            _path = @"\\bbs-actuaries\dfsdata\users\hodsonl\pdfs\test";
            _path += _rnd.Next(1, 1000);
            _path += ".pdf";
        }

        private string GetText(MergeField field)
        {
            var text = "";
            switch (field)
            {
                case MergeField.Title:
                    text = string.Format("The Baptist Pension Scheme (BPS) – {0}", _dataAccess.GetEmployerName());                    
                    break;
                case MergeField.EstimatedEmployerDebt:
                    text = string.Format("Estimated Employer Debt as at {0} ", _dataAccess.GetCessationDate());
                    break;
                case MergeField.IntroListItemOnePartOne:
                    text = "You do not need to take any action ";                    
                    break;
                case MergeField.IntroListItemOnePartTwo:                    
                    text = "as a result of this document, which is for guidance only.  It provides an estimate of the employer debt that your organisation would need to pay, if it were to exit the defined benefit section of the BPS by paying its employer debt immediately.";
                    break;
                case MergeField.IntroListItemTwo:
                    text = "The BPS and its advisers/administrators accept no liability to any organisation for any actions taken (or not taken) as a result of this estimate, the accompanying guidance notes and FAQs. It is each organisation’s responsibility to ensure that it understands the complex legal position in relation to Employer Debts, taking professional advice as necessary.";
                    break;
                case MergeField.IntroListItemThree:
                    text = "There are a number of reasons why the actual figure in your circumstances could differ significantly from the figure set out below.  Please read the notes carefully in case this applies to you.";
                    break;
                case MergeField.IntroListItemFour:
                    text = "Updated figures will be provided on a monthly basis and will rise and fall over time, depending on how the financial position of the Scheme alters.";
                    break;
                case MergeField.IntroListItemFive:
                    text = "This document is for your information. It is not a demand for payment, and you do not need to take any action:";
                    break;
                case MergeField.IntroListItemSix:                    
                    text = "If your organisation has incurred and/or settled a debt in the last 3 months then this estimate might not reflect your up-to-date position.  This is because the details of each employer’s status that are used in the estimated debt calculation are updated once per calendar quarter, so will not reflect more recent cessation events or debt payments.";            
                    break;
                case MergeField.IntroListItemSeven:
                    text = "The estimate provided in this document might not reflect the total amount that would be due if your organisation incurs a cessation event.  In particular: \n";
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
            EstimatedEmployerDebt,
            IntroListItemOnePartOne,
            IntroListItemOnePartTwo,
            IntroListItemTwo,
            IntroListItemThree,
            IntroListItemFour,
            IntroListItemFive,
            IntroListItemSix,
            IntroListItemSeven
        }
        #endregion
    }
}
