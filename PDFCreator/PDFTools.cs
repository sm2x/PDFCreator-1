using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PDFCreator.DataAccess;

namespace PDFCreator
{
    public class PDFTools
    {
        Document _document;
        float _pageWidth;
        float _pageHeight;
        float _verticalPosition;
        float _titleParaWidth;
        float _titleParaHeight;
        float _horizontalOffset;
        Color _baptistBlue;


        public PDFTools(Document document, float pageWidth, float pageHeight, float verticalPosition, float titleParaWidth, float titleParaHeight, float horizontalOffset, Color baptistBlue)
        {
            _document = document;
            _pageWidth = pageWidth;
            _pageHeight = pageHeight;
            _verticalPosition = verticalPosition;
            _titleParaWidth = titleParaWidth;
            titleParaHeight = _titleParaHeight;
            _horizontalOffset = horizontalOffset;
            _baptistBlue = baptistBlue;
            _document.SetMargins(80, 100, 50, 100);                                               
        }

        /// <summary>
        /// For adding basic tables to a fixed position on a single page
        /// </summary>
        /// <param name="width"></param>
        /// <param name="padding"></param>
        /// <param name="alignment"></param>
        /// <param name="fontSize"></param>
        /// <param name="tblCells"></param>
        /// <param name="tableColumns"></param>
        /// <param name=""></param>
        public void AddTable(float width, float padding, TextAlignment alignment, int fontSize, List<TableCellsDTO> tblCells, float[] tableColumns, int? cellHeight = null)
        {
            Table table1 = new Table(tableColumns);
            table1.SetWidth(width);
            table1.SetPadding(padding);
            table1.SetTextAlignment(alignment);
            table1.SetFontSize(fontSize);
            table1.SetKeepTogether(false);

            foreach (var cell in tblCells)
            {
                if (cell.isBold == null || !cell.isBold.Value)
                {
                    table1.AddCell(new Cell().Add(new Paragraph(cell.CellText)).SetWidth(cell.Width));
                }
                else
                {
                    table1.AddCell(new Cell().Add(new Paragraph(cell.CellText)).SetWidth(cell.Width).SetBold());
                }
            }
            table1.SetFixedPosition(100, _verticalPosition, table1.GetWidth());
            _document.Add(table1);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="width"></param>
        /// <param name="padding"></param>
        /// <param name="alignment"></param>
        /// <param name="fontSize"></param>
        /// <param name="tblCells"></param>
        /// <param name="tableColumns"></param>
        /// <param name="isMultiPage"></param>
        public void AddTable(float width, float padding, TextAlignment alignment, int fontSize, List<TableCellsDTO> tblCells, float[] tableColumns, int cellCount, int marginTop)
        {
            Table table1 = new Table(tableColumns);
            table1.SetWidth(width);
            table1.SetPadding(padding);
            table1.SetTextAlignment(alignment);
            table1.SetFontSize(fontSize);
            table1.SetKeepTogether(false);

            foreach (var cell in tblCells.Skip(cellCount))
            {
                table1.AddCell(new Cell().Add(new Paragraph(cell.CellText)).SetWidth(cell.Width).SetHeight(30));
            }
            
            table1.SetRelativePosition(-20, marginTop, 0, 50);
            _document.Add(table1);
        }

        public void NewPage()
        {
            _document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
        }

        /// <summary>
        /// For adding a regular paragraph
        /// </summary>
        /// <param name="text"></param>
        public void AddParagraph(string text)
        {
            Paragraph para = new Paragraph(text);
            para.SetFontSize(9);
            para.SetTextAlignment(TextAlignment.LEFT);
            para.SetWidth(_pageWidth);
            para.SetSpacingRatio(20);
            para.SetHeight(126);
            para.SetFixedPosition(60, _verticalPosition, _titleParaWidth - 120);
            _document.Add(para);
        }


        /// <summary>
        /// For adding a custom styled paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        /// <param name="alignment"></param>
        /// <param name="isUnderlined"></param>
        /// <param name="isBold"></param>
        /// <param name="leftIndent"></param>
        public void AddParagraph(string text, int fontSize, TextAlignment alignment, bool isUnderlined, bool isBold, float leftIndent)
        {
            Paragraph para = new Paragraph(text);
            para.SetFontSize(fontSize);
            para.SetTextAlignment(alignment);
            para.SetWidth(_pageWidth);
            para.SetSpacingRatio(20);
            if (isUnderlined)
            {
                para.SetUnderline();
            }
            if (isBold)
            {
                para.SetBold();
            }
            para.SetHeight(126);
            para.SetFixedPosition(leftIndent, _verticalPosition, _titleParaWidth - 50);
            _document.Add(para);
        }

        public void AddFooter(int pageNumber)
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

        public void AddLogo()
        {
            string strFilePath = @"\\bbs-actuaries\dfsdata\users\hodsonl\Visual Studio 2017\Projects\PDFCreator\PDFCreator\Images\baptist.png";
            AddImage(strFilePath, 39, 110, (_pageWidth - 180), (_pageHeight - 65));
        }

        public void AddImage(string filePath, int height, int width, float positionX, float positionY)
        {
            ImageData l = ImageDataFactory.Create(filePath);
            Image img = new Image(l);
            img.SetHeight(height);
            img.SetWidth(width);
            img.SetFixedPosition(positionX, positionY);
            _document.Add(img);
        }

        public void AddTitle(string titleText)
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

        /// <summary>
        /// For adding a numbered list with a sub list in some of the items
        /// </summary>
        /// <param name="listItems"></param>
        /// <param name="subListItems"></param>
        /// <param name="padding"></param>
        /// <param name="height"></param>
        /// <param name="subOffset"></param>
        /// <param name="addBorders"></param>
        /// <param name="customSymbol"></param>
        public void AddNumberedList(List<ListItem> listItems, Dictionary<int, List<string>> subListItems, List<int> padding, int height, List<int> subOffset, bool addBorders, List<string> customSymbol)
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
                    subList.SetFixedPosition((_pageWidth - _titleParaWidth - 50), _verticalPosition + subOffset[i], _titleParaWidth);
                    _document.Add(subList);
                }
            }

            //add to page
            if (addBorders)
            {
                numberedList.SetBorder(Border.NO_BORDER).SetBorderBottom(new SolidBorder(1f)).SetBorderTop(new SolidBorder(1f)).SetBorderLeft(new SolidBorder(1f));
            }

            numberedList.SetFixedPosition((_pageWidth - _titleParaWidth - (_horizontalOffset / 2)), _verticalPosition, _titleParaWidth);
            _document.Add(numberedList);
        }

        /// <summary>
        /// For adding a regular numbered list
        /// </summary>
        /// <param name="listItems"></param>
        /// <param name="padding"></param>
        /// <param name="height"></param>
        public void AddNumberedList(List<ListItem> listItems, List<int> padding, int height)
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

        public void MoveDown(int numb)
        {
            _verticalPosition -= numb;
        }

        public void ResetPosition()
        {
            _verticalPosition = _pageHeight;
        }

       
    }
}

