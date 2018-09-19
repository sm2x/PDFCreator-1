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


        public PDFGenerator()
        {
            PageSize ps = PageSize.A4;
            _pageWidth = ps.GetWidth();
            _pageHeight = ps.GetHeight();
            _horizontalOffset = 130;
            _baptistBlue = new DeviceRgb(0, 176, 240);
            _dataAccess = new DataAccess(false);
            _verticalPosition = _pageHeight;
            GenerateFilePath();
        }

        public void SpoolPDF()
        {
            _writer = new PdfWriter(_path);
            _pdf = new PdfDocument(_writer);
            _document = new Document(_pdf);
            _document.SetMargins(0, 100, 0, 100);

            SpoolPageOne();
            SpoolPageTwo();
            SpoolPageThree();
            SpoolPageFour();

            _document.Close();
        }

        private void SpoolPageOne()
        {
            AddLogo();
            MoveDown(220);
            AddTitle(GetText(MergeField.Title));

            MoveDown(30);
            AddTitle(GetText(MergeField.EstimatedEmployerDebtAt));

            MoveDown(40);
            AddTitle("Introduction");

            //List of items
            List introList = new List(ListNumberingType.DECIMAL);
            _titleParaWidth = _pageWidth - _horizontalOffset;
            introList.SetFontSize(10);
            introList.SetTextAlignment(TextAlignment.LEFT);
            introList.SetWidth(_titleParaWidth);
            introList.SetMinHeight(150);
            MoveDown(300);

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
            item5SubList.SetFixedPosition((_pageWidth - _titleParaWidth - 50), _verticalPosition + 115, _titleParaWidth);
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

            introList.SetFixedPosition((_pageWidth - _titleParaWidth - (_horizontalOffset / 2)), _verticalPosition, _titleParaWidth);
            _document.Add(introList);

            //nested list
            List item7SubList = new List(ListNumberingType.ZAPF_DINGBATS_2);
            string previousCEWords = _dataAccess.GetPreviousCEWord();
            item7SubList.Add(string.Format("The total due would be greater (potentially significantly greater) if your organisation has previously incurred a separate cessation event. {0}", previousCEWords));
            item7SubList.Add("The amount due could be higher or lower if your organisation is currently in a “period of grace” (a “period of grace” applies if you have had a cessation event and you do not currently employ an active member, but you have confirmed that you expect to take on a new employee who will become a BPS member soon).  As a period of grace can only be granted if an employer requests it, your organisation should be aware if a period of grace applies in your case.  If an employer in a period of grace does not take on a new active member before the end of a period of grace, that employer’s debt will be calculated based on the finances of the Scheme at the date of the cessation event rather than at a current date.");
            if (string.IsNullOrEmpty(previousCEWords))
            {
                MoveDown(140);
            }
            else
            {
                MoveDown(160);
            }

            item7SubList.SetFontSize(10);
            item7SubList.SetFixedPosition((_pageWidth - _titleParaWidth - 50), _verticalPosition, _titleParaWidth);
            _document.Add(item7SubList);

            AddFooter(1);
        }

        private void SpoolPageTwo()
        {
            //reset position tracker            
            ResetPosition();
            NewPage();
            AddLogo();

            //employer debt list            
            MoveDown(220);

            AddTitle("Estimated Employer Debt");

            MoveDown(-10);

            List<ListItem> debtList = new List<ListItem>();
            ListItem employerDebtItem1 = new ListItem(GetText(MergeField.EstimatedEmployerDebtParagraph));
            debtList.Add(employerDebtItem1);

            List<int> singleItemPadding = new List<int>
            {
                10
            };
            AddNumberedListNormal(debtList, singleItemPadding, 90);

            MoveDown(120);

            //comparison with previous figure list            
            AddTitle("Comparison with previous figure");

            MoveDown(-40);

            List<ListItem> comparisonList = new List<ListItem>();
            ListItem comparisonItem1 = new ListItem(GetText(MergeField.ComparisonPreviousFigure));
            comparisonList.Add(comparisonItem1);
            AddNumberedListNormal(comparisonList, singleItemPadding, 60);

            MoveDown(120);

            //do I need to do anything            
            AddTitle("Do I need to do anything?");

            //add list
            MoveDown(230);

            List<int> padding = new List<int>
            {
                10,
                120,
                10
            };
            List<ListItem> doINeedListItems = new List<ListItem>
            {
                new ListItem(GetText(MergeField.DoINeedItemOne)),
                new ListItem(GetText(MergeField.DoINeedItemTwo)),
                new ListItem(GetText(MergeField.DoINeedItemThree))
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
                120, 30
            };

            AddNumberedListWithSub(doINeedListItems, item2SubList, padding, 330, subOffsets, false, new List<string>());

            MoveDown(110);

            AddTitle("How has the estimated employer debt been calculated?");

            MoveDown(240);

            List<ListItem> employerDebtListItems = new List<ListItem>
            {
                new ListItem(GetText(MergeField.HowDebtCalculatedItemOne)),
                new ListItem(GetText(MergeField.HowDebtCalculatedItemTwo))
            };
            List<int> debtPadding = new List<int>
            {
                10,
                10
            };

            AddNumberedListNormal(employerDebtListItems, debtPadding, 330);

            AddFooter(2);
        }

        private void SpoolPageThree()
        {
            ResetPosition();
            NewPage();
            AddLogo();

            MoveDown(430);

            //AddTitle("Test");

            List<ListItem> employerDebtListItems = new List<ListItem>
            {
                new ListItem(GetText(MergeField.HowDebtCalculatedItemThree)),
                new ListItem(GetText(MergeField.HowDebtCalculatedItemFour)),
                new ListItem(GetText(MergeField.HowDebtCalculatedItemFive)),
            };
            List<int> debtPadding = new List<int>
            {
                10,
                140,
                10
            };
            List<int> subOffsets = new List<int>
            {
                140
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
            AddNumberedListWithSub(employerDebtListItems, howDebtCalculatedSubList, debtPadding, 350, subOffsets, false, customSymbol);

            MoveDown(70);

            AddTitle("How does this relate to the contributions we pay each month?");

            MoveDown(230);

            List<ListItem> howDoesRelateListItems = new List<ListItem>
            {
                new ListItem(GetText(MergeField.HowDoesRelateItemOne)),
                new ListItem(GetText(MergeField.HowDoesRelateItemTwo)),
                new ListItem(GetText(MergeField.HowDoesRelateItemThree)),
            };
            List<int> howDoesPadding = new List<int>
            {
                10,
                10,
                10
            };

            AddNumberedListNormal(howDoesRelateListItems, howDoesPadding, 330);

            AddFooter(3);
        }

        private void SpoolPageFour()
        {
            ResetPosition();
            NewPage();            
            AddLogo();

            MoveDown(230);

            AddTitle("Summary of estimated employer debt calculation");

            MoveDown(10);

            

            AddFooter(4);
        }

        #region helper methods      
        private void GenerateFilePath()
        {
            _path = @"\\bbs-actuaries\dfsdata\users\hodsonl\pdfs\test";
            _path += _rnd.Next(1, 100000);
            _path += ".pdf";
        }

        private void NewPage()
        {
            _document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
        }

        private void AddParagraph(string text)
        {
            Paragraph para = new Paragraph(text);            
            para.SetFontSize(9);            
            para.SetWidth(_pageWidth);
            para.SetHeight(126);
            para.SetFixedPosition((_pageWidth - _titleParaWidth - (_horizontalOffset / 2)), _verticalPosition, _titleParaWidth);            
            _document.Add(para);
        }

        private void AddFooter(int pageNumber)
        {
            Paragraph footer = new Paragraph(string.Format("{0} \n \n The Baptist Pension Trust Limited (A Company Limited by Guarantee. Registered in England No 03481942)", pageNumber));
            _titleParaHeight = 126;
            _titleParaWidth = _pageWidth;
            footer.SetFontSize(8);
            PdfFont font = PdfFontFactory.CreateFont(FontConstants.HELVETICA);
            footer.SetFont(font);
            footer.SetFontColor(new DeviceRgb(128, 128, 128));
            footer.SetTextAlignment(TextAlignment.CENTER);
            footer.SetWidth(_titleParaWidth);
            footer.SetHeight(_titleParaHeight);
            footer.SetFixedPosition((_pageWidth - _titleParaWidth), -70, _titleParaWidth);
            _document.Add(footer);
        }

        private void AddLogo()
        {
            //logo            
            string strFilePath = @"\\bbs-actuaries\dfsdata\users\hodsonl\Visual Studio 2017\Projects\PDFCreator\PDFCreator\Images\baptist.png";
            ImageData l = ImageDataFactory.Create(strFilePath);
            Image img = new Image(l);
            img.SetHeight(39);
            img.SetWidth(110);
            img.SetFixedPosition(_pageWidth - 180, _pageHeight - 65);
            _document.Add(img);
        }

        private void AddTitle(string titleText)
        {
            Paragraph title = new Paragraph(titleText);
            _titleParaHeight = 126;
            _titleParaWidth = _pageWidth;
            title.SetFontSize(13);
            title.SetBold();
            title.SetTextAlignment(TextAlignment.CENTER);
            title.SetFontColor(_baptistBlue);
            title.SetUnderline();
            title.SetWidth(_titleParaWidth);
            title.SetHeight(_titleParaHeight);
            title.SetFixedPosition((_pageWidth - _titleParaWidth), _verticalPosition, _titleParaWidth);
            _document.Add(title);
        }

        private void AddNumberedListWithSub(List<ListItem> listItems, Dictionary<int, List<string>> subListItems, List<int> padding, int height, List<int> subOffset, bool addBorders, List<string> customSymbol)
        {
            //add list
            List numberedList = new List(ListNumberingType.DECIMAL);
            _titleParaWidth = _pageWidth - _horizontalOffset;
            numberedList.SetFontSize(10);
            numberedList.SetTextAlignment(TextAlignment.LEFT);
            numberedList.SetWidth(_titleParaWidth);
            numberedList.SetHeight(height);

            //add the list items
            for (int i = 0; i < listItems.Count; i++)
            {
                listItems[i].SetPaddingBottom(padding[i]);
                numberedList.Add(listItems[i]);

                //add customer symbol if one needed
                if (customSymbol.Count > 0 && !string.IsNullOrEmpty(customSymbol[i]))
                {
                    listItems[i].SetListSymbol(customSymbol[i]);
                }

                //add sub items
                if (subListItems.ContainsKey(i + 1))
                {
                    List subList = new List(ListNumberingType.ZAPF_DINGBATS_2);
                    var subItemList = subListItems[i + 1];
                    foreach (var subItem in subItemList)
                    {
                        subList.Add(subItem);
                    }

                    subList.SetFontSize(10);
                    subList.SetFixedPosition((_pageWidth - _titleParaWidth - 50), _verticalPosition + subOffset[i - 1], _titleParaWidth);
                    _document.Add(subList);
                }
            }

            //add to page
            if (addBorders)
            {
                numberedList.SetBorder(Border.NO_BORDER).SetBorderBottom(new SolidBorder(1f)).SetBorderTop(new SolidBorder(1f)).SetBorderLeft(new SolidBorder(1f));
            }

            numberedList.SetKeepTogether(true);
            numberedList.SetFixedPosition((_pageWidth - _titleParaWidth - (_horizontalOffset / 2)), _verticalPosition, _titleParaWidth);
            _document.Add(numberedList);
        }

        private void AddNumberedListNormal(List<ListItem> listItems, List<int> padding, int height)
        {
            //add list
            List numberedList = new List(ListNumberingType.DECIMAL);
            _titleParaWidth = _pageWidth - _horizontalOffset;
            numberedList.SetFontSize(10);
            numberedList.SetTextAlignment(TextAlignment.LEFT);
            numberedList.SetWidth(_titleParaWidth);
            numberedList.SetHeight(height);

            //add the list items
            for (int i = 0; i < listItems.Count; i++)
            {
                listItems[i].SetPaddingBottom(padding[i]);
                numberedList.Add(listItems[i]);
            }

            //add to page            
            numberedList.SetFixedPosition((_pageWidth - _titleParaWidth - (_horizontalOffset / 2)), _verticalPosition, _titleParaWidth);
            _document.Add(numberedList);
        }

        private void MoveDown(int numb)
        {
            _verticalPosition -= numb;
        }

        private void ResetPosition()
        {
            _verticalPosition = _pageHeight;
        }

        private enum MergeField
        {
            Title,
            EstimatedEmployerDebtAt,
            IntroListItemOnePartOne,
            IntroListItemOnePartTwo,
            IntroListItemTwo,
            IntroListItemThree,
            IntroListItemFour,
            IntroListItemFive,
            IntroListItemSix,
            IntroListItemSeven,
            EstimatedEmployerDebtParagraph,
            ComparisonPreviousFigure,
            DoINeedItemOne,
            DoINeedItemTwo,
            DoINeedItemThree,
            HowDebtCalculatedItemOne,
            HowDebtCalculatedItemTwo,
            HowDebtCalculatedItemThree,
            HowDebtCalculatedItemFour,
            HowDebtCalculatedItemFive,
            HowDoesRelateItemOne,
            HowDoesRelateItemTwo,
            HowDoesRelateItemThree
        }


        private string GetText(MergeField field)
        {
            var text = "";
            switch (field)
            {
                case MergeField.Title:
                    text = string.Format("The Baptist Pension Scheme (BPS) – {0}", _dataAccess.GetEmployerName());
                    break;
                case MergeField.EstimatedEmployerDebtAt:
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
                case MergeField.EstimatedEmployerDebtParagraph:
                    text = string.Format("The estimated employer debt for your organisation at {0} is {1}.  The calculation is summarised in Table 1 below. This figure has been calculated using this “assumed” cessation date based on financial conditions near the month end, but unless your organisation happened to have an actual cessation event on this date, no employer debt is due to be paid at this time based on this cessation date and this estimated debt.  The figure is for information only.", _dataAccess.GetCessationDate(), _dataAccess.GetCessationAmount());
                    break;
                case MergeField.ComparisonPreviousFigure:
                    text = "Figures for individual employers may vary from month to month as a result of changes in Scheme membership (for example, retirements, deaths or transfers out of the Scheme), as well as reflecting the general trend.";
                    break;
                case MergeField.DoINeedItemOne:
                    text = "This document is not a demand for payment, and you do not need to take any action.  Your organisation can simply use this information for monitoring the changes to the estimated employer debt over time, while continuing to make the required monthly deficit contributions, provided it continues to employ an active member of the BPS.";
                    break;
                case MergeField.DoINeedItemTwo:
                    text = "If your organisation were to consider incurring a cessation event and settling its employer debt, you should note the following:";
                    break;
                case MergeField.DoINeedItemThree:
                    text = "To understand fully the implications of and timescales for any decision, you must also read the accompanying documents:";
                    break;
                case MergeField.HowDebtCalculatedItemOne:
                    text = "The estimated employer debt is calculated based on the BPS’ funding position and the Trustee's knowledge of the BPS’ liabilities as at the assumed cessation date, based on financial conditions near the month end.  More details are provided in the Guidance Notes and Frequently Asked Questions, and a summary of the calculation for your organisation is shown in Table 1.";
                    break;
                case MergeField.HowDebtCalculatedItemTwo:
                    text = "The estimated employer debt is calculated using the Trustee's record of your organisation’s current and former ministers who were in service with you, as set out in Table 2 of this document. If you do not agree with the record in Table 2, please let us know by emailing Mark Hynes, the Pensions Manager on mhynes@baptist.org.uk";
                    break;
                case MergeField.HowDebtCalculatedItemThree:
                    text = "Your organisation’s pension information is confidential to the BPS and cannot be shared without your permission. However, the Baptist Union Regional Associations have supported a number of churches and other scheme employers understand their obligations and options and you may find it helpful to contact them.";
                    break;
                case MergeField.HowDebtCalculatedItemFour:
                    text = "The size of an employer’s liability to the BPS depends on two main factors:";
                    break;
                case MergeField.HowDebtCalculatedItemFive:
                    text = "It is important to note that the Employer Debt Regulations (2005 and as amended), require any liabilities which cannot specifically be attributed to any current employer (“orphan liabilities”) to be shared amongst all current employers.  These orphan liabilities are part of the Scheme deficit calculation.";
                    break;
                case MergeField.HowDoesRelateItemOne:
                    text = "Your monthly deficiency payments to the defined benefit plan are set every three years to target the deficit in the Scheme at the time.  This deficit is measured using different assumptions from those used to calculate the estimated employer debt.  As a result the monthly contributions are targeting a lower deficit than the very prudent measure that the regulations say must be used for employer debt calculations.";
                    break;
                case MergeField.HowDoesRelateItemTwo:
                    text = "The monthly deficiency payments are at present similar for most employers.  This contrasts with the employer debt calculations, which depend directly on the liabilities in the BPS that relate to each employer.";
                    break;
                case MergeField.HowDoesRelateItemThree:
                    text = "As a result, your estimated employer debt could look very large, or indeed quite small, compared with the amount you might expect to pay on a monthly basis over the period to 2028 (which is the end point of the current plan to address the Scheme deficit).";
                    break;
                default:
                    break;
            }
            return text;
        }
        #endregion
    }
}
