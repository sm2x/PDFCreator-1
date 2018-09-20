using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Colorspace;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PDFCreator.DataAccess;

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
        float _horizontalOffset;
        DataAccess _dataAccess;
        float _verticalPosition;
        float _titleParaHeight;
        float _titleParaWidth;
        PDFTools tls;


        public PDFGenerator()
        {
            PageSize ps = PageSize.A4;
            _pageWidth = ps.GetWidth();
            _pageHeight = ps.GetHeight();
            _horizontalOffset = 130;
            _baptistBlue = new DeviceRgb(0, 176, 240);
            _dataAccess = new DataAccess(true);
            _verticalPosition = _pageHeight;

            GenerateFilePath(); //remove
            _writer = new PdfWriter(_path);
            _pdf = new PdfDocument(_writer);
            _document = new Document(_pdf);
            tls = new PDFTools(_document, _pageWidth, _pageHeight, _verticalPosition, _titleParaWidth, _titleParaHeight, _horizontalOffset, _baptistBlue);
        }

        public void SpoolPDF()
        {
            SpoolPageOne();
            SpoolPageTwo();
            SpoolPageThree();
            SpoolPageFour();
            SpoolPageFive();
            _document.Close();
        }

        private void SpoolPageOne()
        {
            tls.AddLogo();
            tls.MoveDown(220);
            tls.AddTitle(_dataAccess.GetText(MergeField.Title));

            tls.MoveDown(30);
            tls.AddTitle(_dataAccess.GetText(MergeField.EstimatedEmployerDebtAt));

            tls.MoveDown(40);
            tls.AddTitle("Introduction");

            tls.MoveDown(630);   

            //Item 1
            Paragraph para = new Paragraph();
            para.Add(new Text(_dataAccess.GetText(MergeField.IntroListItemOnePartOne)).SetBold());
            para.Add(new Text(_dataAccess.GetText(MergeField.IntroListItemOnePartTwo)));

            ListItem listItemOne = new ListItem();
            listItemOne.Add(para);
            listItemOne.SetPaddingBottom(10);

            //Item 2            
            ListItem listItemTwo = new ListItem(_dataAccess.GetText(MergeField.IntroListItemTwo));
            listItemTwo.SetPaddingBottom(10);

            //Item 3  
            ListItem listItemThree = new ListItem(_dataAccess.GetText(MergeField.IntroListItemThree));
            listItemThree.SetPaddingBottom(10);

            //Item 4
            ListItem listItemFour = new ListItem(_dataAccess.GetText(MergeField.IntroListItemFour));
            listItemFour.SetPaddingBottom(10);

            //Item 5
            ListItem listItemFive = new ListItem(_dataAccess.GetText(MergeField.IntroListItemFive));
            listItemFive.SetPaddingBottom(10);

            //Item 6
            ListItem listItemSix = new ListItem(_dataAccess.GetText(MergeField.IntroListItemSix));
            listItemSix.SetPaddingBottom(10);
            listItemSix.SetPaddingTop(40);

            //item 7
            ListItem listItemSeven = new ListItem(_dataAccess.GetText(MergeField.IntroListItemSeven));
            listItemSeven.SetPaddingBottom(10);

            List<ListItem> introListItems = new List<ListItem>
            {
                listItemOne,
                listItemTwo,
                listItemThree,
                listItemFour,
                listItemFive,
                listItemSix,
                listItemSeven
            };

            //add sub list
            string previousCEWords = _dataAccess.GetPreviousCEWord();
            Dictionary<int, List<string>> introSubList = new Dictionary<int, List<string>>
            {
                {   5, //list item 2
                    new List<string>
                    {
                        "an employer debt only becomes due when an employer incurs a “cessation event”",
                        "a cessation event normally only occurs when an employer stops employing any active members of the BPS"
                    }
                },
                {
                    7,
                    new List<string>
                    {
                        string.Format("The total due would be greater (potentially significantly greater) if your organisation has previously incurred a separate cessation event. {0}", previousCEWords),
                        "The amount due could be higher or lower if your organisation is currently in a “period of grace” (a “period of grace” applies if you have had a cessation event and you do not currently employ an active member, but you have confirmed that you expect to take on a new employee who will become a BPS member soon).  As a period of grace can only be granted if an employer requests it, your organisation should be aware if a period of grace applies in your case.  If an employer in a period of grace does not take on a new active member before the end of a period of grace, that employer’s debt will be calculated based on the finances of the Scheme at the date of the cessation event rather than at a current date.",
                    }
                }
            };

            List<int> subOffsets = new List<int>
            {
                0, 0, 0, 0, 440, 0, 170
            };

            //subliststart
            List<int> intoListPadding = new List<int>
            {
                10,
                10,
                10,
                10,
                string.IsNullOrEmpty(previousCEWords) ? 20 : 10,
                10,
                10
            };

            tls.AddNumberedListWithSub(introListItems, introSubList, intoListPadding, 730, subOffsets, false, new List<string>());          

            tls.AddFooter(1);
        }

        private void SpoolPageTwo()
        {
            //reset position tracker            
            tls.ResetPosition();
            tls.NewPage();
            tls.AddLogo();

            //employer debt list            
            tls.MoveDown(220);

            tls.AddTitle("Estimated Employer Debt");

            tls.MoveDown(-10);

            List<ListItem> debtList = new List<ListItem>();
            ListItem employerDebtItem1 = new ListItem(_dataAccess.GetText(MergeField.EstimatedEmployerDebtParagraph));
            debtList.Add(employerDebtItem1);

            List<int> singleItemPadding = new List<int>
            {
                10
            };
            tls.AddNumberedListNormal(debtList, singleItemPadding, 90);

            tls.MoveDown(120);

            //comparison with previous figure list            
            tls.AddTitle("Comparison with previous figure");

            tls.MoveDown(-40);

            List<ListItem> comparisonList = new List<ListItem>();
            ListItem comparisonItem1 = new ListItem(_dataAccess.GetText(MergeField.ComparisonPreviousFigure));
            comparisonList.Add(comparisonItem1);
            tls.AddNumberedListNormal(comparisonList, singleItemPadding, 60);

            tls.MoveDown(120);

            //do I need to do anything            
            tls.AddTitle("Do I need to do anything?");

            //add list
            tls.MoveDown(230);

            List<int> padding = new List<int>
            {
                10,
                120,
                10
            };
            List<ListItem> doINeedListItems = new List<ListItem>
            {
                new ListItem(_dataAccess.GetText(MergeField.DoINeedItemOne)),
                new ListItem(_dataAccess.GetText(MergeField.DoINeedItemTwo)),
                new ListItem(_dataAccess.GetText(MergeField.DoINeedItemThree))
            };

            //add sub list
            Dictionary<int, List<string>> item2SubList = new Dictionary<int, List<string>>
            {
                {   2, //list item 2
                    new List<string>
                    {
                        "The estimated figure provided in this note is only valid for the calendar month following the date of calculation, after which it will be updated to show the latest position",
                        "The process for settling an employer debt is complex and must be completed within very tight timescales.",
                        "Settlement on the basis of an estimated employer debt can only take place with the agreement of the pension Trustee. There may be circumstances in which the Trustee requires the debt to be certified instead."
                    }
                },
                {
                     3,
                    new List<string>
                    {
                        "Automated Monthly Debt Estimate – Employer Guidance Notes",
                        "Employer Debts – Frequently Asked Questions",
                        "Automated Monthly Debt Estimate – Settlement Process Guide"
                    }
                }
            };
            List<int> subOffsets = new List<int>
            {
                0, 120, 30
            };

            tls.AddNumberedListWithSub(doINeedListItems, item2SubList, padding, 330, subOffsets, false, new List<string>());

            tls.MoveDown(110);

            tls.AddTitle("How has the estimated employer debt been calculated?");

            tls.MoveDown(240);

            List<ListItem> employerDebtListItems = new List<ListItem>
            {
                new ListItem(_dataAccess.GetText(MergeField.HowDebtCalculatedItemOne)),
                new ListItem(_dataAccess.GetText(MergeField.HowDebtCalculatedItemTwo))
            };
            List<int> debtPadding = new List<int>
            {
                10,
                10
            };

            tls.AddNumberedListNormal(employerDebtListItems, debtPadding, 330);

            tls.AddFooter(2);
        }

        private void SpoolPageThree()
        {
            tls.ResetPosition();
            tls.NewPage();
            tls.AddLogo();

            tls.MoveDown(430);

            //AddTitle("Test");

            List<ListItem> employerDebtListItems = new List<ListItem>
            {
                new ListItem(_dataAccess.GetText(MergeField.HowDebtCalculatedItemThree)),
                new ListItem(_dataAccess.GetText(MergeField.HowDebtCalculatedItemFour)),
                new ListItem(_dataAccess.GetText(MergeField.HowDebtCalculatedItemFive)),
            };
            List<int> debtPadding = new List<int>
            {
                10,
                140,
                10
            };
            List<int> subOffsets = new List<int>
            {
               0, 140
            };
            Dictionary<int, List<string>> howDebtCalculatedSubList = new Dictionary<int, List<string>>
            {
                {   2, //list item 2
                    new List<string>
                    {
                        "The estimated deficit in the Scheme at the assumed cessation date.  For the employer debt calculations, this deficit is the difference between the estimated cost of securing all members’ defined benefits with an insurance company and the value of the Scheme’s total assets held.",
                        "Your organisation’s share of the deficit.  This depends on the value of the benefits earned by your organisation’s current and former ministers who were members of the defined benefit plan whilst in service with you (and in certain circumstances, while in ministry prior to joining your organisation), including any supplementary benefits accrued by those ministers.  It also depends on the equivalent figures for all the other employers still participating in the Scheme."
                    }
                }
            };
            //when using custom symbols, must use custom symbols for all the list
            List<string> customSymbol = new List<string>
            {
                "3. ",
                "4. ",
                "5. "
            };
            tls.AddNumberedListWithSub(employerDebtListItems, howDebtCalculatedSubList, debtPadding, 350, subOffsets, false, customSymbol);

            tls.MoveDown(70);

            tls.AddTitle("How does this relate to the contributions we pay each month?");

            tls.MoveDown(230);

            List<ListItem> howDoesRelateListItems = new List<ListItem>
            {
                new ListItem(_dataAccess.GetText(MergeField.HowDoesRelateItemOne)),
                new ListItem(_dataAccess.GetText(MergeField.HowDoesRelateItemTwo)),
                new ListItem(_dataAccess.GetText(MergeField.HowDoesRelateItemThree)),
            };
            List<int> howDoesPadding = new List<int>
            {
                10,
                10,
                10
            };

            tls.AddNumberedListNormal(howDoesRelateListItems, howDoesPadding, 330);

            tls.AddFooter(3);
        }

        private void SpoolPageFour()
        {
            tls.ResetPosition();
            tls.NewPage();
            tls.AddLogo();

            tls.MoveDown(230);

            tls.AddTitle("Summary of estimated employer debt calculation");

            tls.MoveDown(30);

            tls.AddParagraph("The calculation is set out in Table 1 below and can be summarised as:");

            string strFilePath = @"\\bbs-actuaries\dfsdata\users\hodsonl\Visual Studio 2017\Projects\PDFCreator\PDFCreator\Images\formula.png";
            tls.AddImage(strFilePath, 49, 350, _pageHeight - 720, 640);

            tls.MoveDown(100);

            tls.AddCustomParagraph("Table 1", 10, TextAlignment.LEFT, true, true, 100);

            tls.MoveDown(290);
            
            //table column widths
            float leftWidth = 100;
            float centerWidth = 160;
            float rightWidth = 160;

            //add all the table cells
            List<TableCellsDTO> tblCells = new List<TableCellsDTO>
            {
                new TableCellsDTO { CellText =  "", Width = leftWidth },
                new TableCellsDTO { CellText =  "Figures for your organisation", Width = centerWidth, isBold = true },
                new TableCellsDTO { CellText =  "Explanation", Width = rightWidth, isBold = true },

                new TableCellsDTO { CellText =  "Cessation date", Width = leftWidth },
                new TableCellsDTO { CellText =  "«CessationDate»(assumed)", Width = centerWidth },
                new TableCellsDTO { CellText =  "Assumed cessation event for illustration. ", Width = rightWidth },

                new TableCellsDTO { CellText =  "Liability value relating to your organisation(A)", Width = leftWidth },
                new TableCellsDTO { CellText =  "£«ChurchLiability»", Width = centerWidth },
                new TableCellsDTO { CellText =  "This is the BPS actuary’s estimated cost of securing with an insurance company the pension benefits for your organisation - based on membership data as detailed in Table 2 below - as at the assumed cessation date.", Width = rightWidth },

                new TableCellsDTO { CellText =  "Liability value relating to all current employers(B)", Width = leftWidth },
                new TableCellsDTO { CellText =  "£«TotalAttributableLiability»", Width = centerWidth },
                new TableCellsDTO { CellText =  "This is the BPS actuary’s estimated cost of securing with an insurance company all the Scheme’s pension liabilities relating to current employers (ie excluding “orphan liabilities”), as at the assumed cessation date", Width = rightWidth },

                new TableCellsDTO { CellText =  "Total Scheme deficit(C)", Width = leftWidth },
                new TableCellsDTO { CellText =  "£«TotalDeficit»", Width = centerWidth },
                new TableCellsDTO { CellText =  "This is the BPS actuary’s estimate of the deficit when comparing the Scheme’s total assets with the cost of securing with an insurance company all the Scheme’s pension liabilities (ie including “orphan liabilities”),  as at the assumed cessation date", Width = rightWidth },

                new TableCellsDTO { CellText =  "Cessation Expenses", Width = leftWidth },
                new TableCellsDTO { CellText =  "£«CessationExpenses»", Width = centerWidth },
                new TableCellsDTO { CellText =  "Cessation expenses ", Width = rightWidth },

                new TableCellsDTO { CellText =  "Estimated employer debt ", Width = leftWidth },
                new TableCellsDTO { CellText =  "£«EmployerDebt»", Width = centerWidth },
                new TableCellsDTO { CellText =  "Based on the calculation of (A) / (B) x (C) + Cessation Expenses", Width = rightWidth },
            };

            tls.AddTable(420, 5, TextAlignment.CENTER, 9, tblCells);          
            tls.AddFooter(4);
        }

        private void SpoolPageFive()
        {
            tls.ResetPosition();
            tls.NewPage();
            tls.AddLogo();

            tls.MoveDown(230);

            tls.AddTitle("Summary of membership data for your organisation");



            tls.AddFooter(4);
        }


        public void GenerateFilePath()
        {
            _path = @"\\bbs-actuaries\dfsdata\users\hodsonl\pdfs\test";
            _path += _rnd.Next(1, 100000);
            _path += ".pdf";
        }
    }
}
