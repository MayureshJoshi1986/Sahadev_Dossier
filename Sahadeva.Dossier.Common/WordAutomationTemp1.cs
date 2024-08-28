using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Configuration;
using Sahadeva.Dossier.Entities;
using System.Data;
using System.Net;
using System.IO;
using System.Web;

namespace Sahadeva.Dossier.Common
{
  public  class WordAutomationTemp1
    {
        public Microsoft.Office.Interop.Word.Application oWord;	// a reference to Word application
        public Microsoft.Office.Interop.Word.Document oDoc; // a reference to the document

        object oMissing = System.Reflection.Missing.Value;
        //OBJECTS OF FALSE AND TRUE
        Object oTrue = true;
        Object oFalse = false;
        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
        Color Orangecolor = Color.FromArgb(238, 124, 30);
        Color FadedOrangecolor = Color.FromArgb(255, 222, 173);
        Color Whitecolor = Color.FromArgb(255, 255, 255);
        Color Colomncolor = Color.FromArgb(241, 235, 185);
        Color lightrowcolor = Color.FromArgb(254, 252, 235);
        Color darkrowcolor = Color.FromArgb(243, 240, 216);
        Color LightBlueColor = Color.FromArgb(230, 245, 255);
        Color DarkBlueColor = Color.FromArgb(128, 170, 255);

        public WordAutomationTemp1()
        {

            // activate the interface with the COM object of Microsoft Word
            //oWord = new Microsoft.Office.Interop.Word.Application();
            //oDoc = oWord.Documents.Add(@"D:\Dosier\WordDocWithTemplate\WordDocWithTemplate\Templates\Template.dotx", oMissing, oMissing, oMissing);
            //oWord.Visible = false;

            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add(oMissing, oMissing, oMissing, oMissing);
            //oWord.Visible = false;
            oDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

            //  oDoc.PageSetup.TopMargin=
            oDoc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;

            // oDoc.PageSetup.PageWidth = 827f; // Width for A4
            // oDoc.PageSetup.PageHeight = 1169f; // Height for A4

            // Adjust margins to fit within the page width
            //oDoc.PageSetup.LeftMargin = 100f; // Set the left margin in inches
            //oDoc.PageSetup.RightMargin = 100f; // Set the right margin in inches
            // Adjust other margins if needed (TopMargin, BottomMargin, etc.)









        }
        public WordAutomationTemp1(string WordTemplateFilePath)
        {
            // activate the interface with the COM object of Microsoft Word
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add(@"" + WordTemplateFilePath + "", oMissing, oMissing, oMissing);
            //oWord.Visible = false;
        }
        public void InsertPagebreak()
        {
            // VB : Selection.InsertBreak Type:=wdPageBreak
            object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
            oWord.Selection.InsertBreak(ref pBreak);
            //oWord.ActiveWindow.Selection.InsertBreak(ref pBreak);
            //oWord.ActiveWindow.Selection.InsertAfter("\n");
        }
        public void InserLineBreak()
        {
            object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak;
            oWord.Selection.InsertBreak(ref pBreak);
        }

        public void InsertHeaderandFooter(string HeaderImageUrl, string FooterImageUrl)
        {
            

            //old
            //foreach (Section section in oDoc.Sections)
            //{
            //    section.PageSetup.PageHeight = 1100;
            //    word.HeaderFooter header = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
            //    Shape oheader = header.Shapes.AddPicture(HeaderImageUrl, oMissing, oMissing, -20, -20);
            //    oheader.Left = section.PageSetup.PageWidth - oheader.Width - 330;
            //    oheader.Top = -10;

            //if (FooterImageUrl != "")
            //{
            //    word.HeaderFooter footer = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

            //    // Add page numbers to the footer
            //    Range pageNumberRange = footer.Range;
            //    Field pageField = pageNumberRange.Fields.Add(pageNumberRange, WdFieldType.wdFieldPage);

            //    // Optionally, you can customize the appearance of the page number
            //    pageNumberRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            //    Shape ofooter = footer.Shapes.AddPicture(FooterImageUrl, oMissing, oMissing, -20, -20);
            //    ofooter.Left = section.PageSetup.PageWidth - ofooter.Width - 310;
            //    //ofooter.Top = footer.Range.Information[WdInformation.wdVerticalPositionRelativeToPage];

            //    // Update the fields to reflect the changes
            //    oDoc.Fields.Update();
            //}

               
            //}

            foreach (Section section in oDoc.Sections)
            {
                section.PageSetup.PageHeight = 1100;
                word.HeaderFooter header = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                Shape oheader = header.Shapes.AddPicture(HeaderImageUrl, oMissing, oMissing, -20, -20);
                oheader.Left = section.PageSetup.PageWidth - oheader.Width - 330;//(section.PageSetup.PageWidth - oheader.Width) / 2;  //
                oheader.Top = -10;



                float maxWidth = 210.0f; // Maximum width
                float maxHeight = 50.0f; // Maximum height

                // Check if the current height exceeds the maximum height
                //if (oheader.Height > maxHeight)
                //{
                //    // Calculate the aspect ratio
                //    float aspectRatio = oheader.Width / oheader.Height;

                //    // Set the height to the maximum height
                //    oheader.Height = maxHeight;

                //    // Set the width based on the aspect ratio to maintain the aspect ratio
                //    oheader.Width = oheader.Height * aspectRatio;

                //    // Check if the width exceeds the maximum width
                //    if (oheader.Width > maxWidth)
                //    {
                //        // Set the width to the maximum width
                //        oheader.Width = maxWidth;

                //        // Set the height based on the aspect ratio to maintain the aspect ratio
                //        oheader.Height = oheader.Width / aspectRatio;
                //    }
                //}
                if (FooterImageUrl != "")
                {
                    word.HeaderFooter footer = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    // Add page numbers to the footer
                    Range pageNumberRange = footer.Range;
                    Field pageField = pageNumberRange.Fields.Add(pageNumberRange, WdFieldType.wdFieldPage);

                    // Optionally, you can customize the appearance of the page number
                    pageNumberRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                    Shape ofooter = footer.Shapes.AddPicture(FooterImageUrl, oMissing, oMissing, -20, -20);
                    ofooter.Left = section.PageSetup.PageWidth - ofooter.Width - 310;
                    //ofooter.Top = footer.Range.Information[WdInformation.wdVerticalPositionRelativeToPage];

                    // Update the fields to reflect the changes
                    oDoc.Fields.Update();
                }

                //if (FooterImageUrl != "")
                //{
                //    word.HeaderFooter footer = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

                //    // Add page numbers to the footer
                //    Range pageNumberRange = footer.Range;
                //    Field pageField = pageNumberRange.Fields.Add(pageNumberRange, WdFieldType.wdFieldPage);

                //    // Optionally, you can customize the appearance of the page number
                //    pageNumberRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                //    Shape ofooter = footer.Shapes.AddPicture(FooterImageUrl, oMissing, oMissing, -20, -20);
                //    ofooter.Top = section.PageSetup.PageWidth + 150;
                //    ofooter.Left = 130;
                //    //ofooter.Top = footer.Range.Information[WdInformation.wdVerticalPositionRelativeToPage];

                //    // Update the fields to reflect the changes
                //    oDoc.Fields.Update();
                //}


            }
           
            oDoc.PageSetup.TopMargin = (float)80;
            oDoc.PageSetup.BottomMargin = (float)80;
            oDoc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;

        }

        public void InsertHeaderandFooter_New(string HeaderImageUrl, string FooterImageUrl)
        {
            foreach (Section section in oDoc.Sections)
            {
                section.PageSetup.PageHeight = 1100;
                word.HeaderFooter header = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                Shape oheader = header.Shapes.AddPicture(HeaderImageUrl, oMissing, oMissing, -20, -20);
                oheader.Left = -45;
                oheader.Top = -30;

                Shape oheader1 = header.Shapes.AddPicture(FooterImageUrl, oMissing, oMissing, oMissing, oMissing);
                oheader1.Left = section.PageSetup.PageWidth - oheader1.Width - 90; // Adjust the position as needed
                oheader1.Top = -20;

                oheader.PictureFormat.Contrast = (float)1;
                //oheader1.PictureFormat.Contrast= (float)(0.2);
            }
            oDoc.PageSetup.TopMargin = (float)100;
            //oDoc.PageSetup.BottomMargin = (float)80;

        }

        public void InsertImageInFooter(string ImageUrl)
        {

            // oWord.Selection.TypeText(File.ReadAllText(@"D:\Notepads\XML.txt"));

            // (int)section.Range.get_Information(WdInformation.wdActiveEndSectionNumber)
            //Microsoft.Office.Interop.Word.Shape logoCustomer = null;
            //Microsoft.Office.Interop.Word.Shape logoadf = null;

            foreach (Section section in oDoc.Sections)
            {
                //section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "Page" + section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers.Count;
                //section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InlineShapes.AddPicture(ImageUrl); ;
                section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InlineShapes.AddPicture(ImageUrl);

                section.PageSetup.PageHeight = 1150;



            }


            //// Open up the footer in the word document
            //oDoc.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter;


            //// Set current Paragraph Alignment to Center
            //oDoc.ActiveWindow.ActivePane.Selection.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


            //// Type in 'Page '
            //oDoc.ActiveWindow.Selection.TypeText("Page ");


            //// Add in current page field
            //Object CurrentPage = WdFieldType.wdFieldPage;

            //oDoc.ActiveWindow.Selection.Fields.Add(oDoc.ActiveWindow.Selection.Range,

            //    ref CurrentPage, Type.Missing, Type.Missing);


            //// Type in ' of '
            //oDoc.ActiveWindow.Selection.TypeText(" of ");


            //// Add in total page field
            //Object TotalPages = WdFieldType.wdFieldNumPages;

            //oDoc.ActiveWindow.Selection.Fields.Add(oDoc.ActiveWindow.Selection.Range,

            //    ref TotalPages, Type.Missing, Type.Missing);
        }
        public void InsertText(string Text)
        {
            //word.Application oWord = new word.Application();
            //oWord.Visible = false;
            //object oMissing = System.Reflection.Missing.Value;

            // word.Document oDoc = new word.Document();
            //oDoc = oWord.Documents.Add(oMissing, oMissing, oMissing, oMissing);
            oWord.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oWord.Selection.Font.Bold = 0;
            oWord.Selection.TypeText(Text);
            oWord.ActiveWindow.Selection.InsertAfter("\n");

            //oWord.Visible = true;
            //oWord = null;
        }
        public void InsertText_WithoutNewLine(string Text)
        {
            oWord.Selection.TypeText(Text);
        }

        public void InsertHyperLink(string hyperlinkText, string hyperlinkAddress)
        {
            oWord.Selection.Hyperlinks.Add(oWord.Selection.Range, hyperlinkAddress, oMissing, oMissing, hyperlinkText, oMissing);
        }

        public void InsertDocumentHeadline(string Text)
        {
            Insert_NewLine();

            // Microsoft.Office.Interop.Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);

            // para.Range.InsertParagraphAfter();

            //  para.Range.Text = Text;
            // Explicitly set this to "not bold"
            // para.Range.Font.Bold = 7;
            // para.Range.Font.Size = 15;

            //  para.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //para.Format.SpaceAfter = piSpaceAfter;

            //object objStart = para.Range.Start;
            //object objEnd = para.Range.Start + psText.IndexOf(":");

            //Word.Range rngBold = mdocWord.Range(ref objStart, ref objEnd);
            //rngBold.Bold = 1;

            oWord.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

            oWord.Selection.TypeText(Text);
            oWord.Selection.Font.Bold = 7;
            oWord.Selection.Font.Size = 15;

        }

        public void InsertHeaderTable(string Publication, string Edition, string Date, string Page)
        {
            //Microsoft.Office.Interop.Word._Application objWord;
            //Microsoft.Office.Interop.Word._Document objDoc;
            //objWord = new Microsoft.Office.Interop.Word.Application();
            //objWord.Visible = true;
            //objDoc = objWord.Documents.Add(ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing);

            int i = 0;
            int j = 0;
            Microsoft.Office.Interop.Word.Table objTable;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            objTable = oDoc.Tables.Add(wrdRng, 2, 2, ref oMissing, ref oMissing);
            objTable.Range.ParagraphFormat.SpaceAfter = 2;
            objTable.Rows.Alignment = WdRowAlignment.wdAlignRowLeft;

            objTable.Cell(1, 1).Range.Text = "Publication : " + Publication;
            objTable.Cell(1, 2).Range.Text = "Edition : " + Edition;
            objTable.Cell(2, 1).Range.Text = "Date : " + Date;
            objTable.Cell(2, 2).Range.Text = "Page: " + Page;

            objTable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            objTable.Rows[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;


            objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
            objTable.Rows[2].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Colomncolor);

            objTable.Rows[1].Range.Borders.Enable = 1;
            objTable.Rows[2].Range.Borders.Enable = 1;
            objTable.Columns[1].Borders.Enable = 1;


            objTable.Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

            objTable.Rows[1].Range.Font.Bold = 7;
            objTable.Rows[1].Range.Font.ColorIndex = WdColorIndex.wdWhite;

            objTable.Columns[1].Borders.OutsideColor = WdColor.wdColorWhite;
            objTable.Columns[2].Borders.OutsideColor = WdColor.wdColorWhite;


            objTable.Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

            //objTable.Rows[1].Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
            //objTable.Rows[1].Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

            //objTable.BottomPadding = (float)0.5;
            //objTable.TopPadding = (float)0.5;
            //oWord.ActiveWindow.Selection.InsertAfter("\n");

        }

        public void InsertHeaderTable_New_Print(NewsArticle objNewsArticle)
        {

            try
            {

                int i = 0;
                int j = 0;
                Microsoft.Office.Interop.Word.Table objTable;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                objTable = oDoc.Tables.Add(wrdRng, 7, 5, ref oMissing, ref oMissing);
                objTable.Range.ParagraphFormat.SpaceAfter = 1;
                objTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                // Set column widths (adjust as needed)
                objTable.Columns[1].Width = 100;
                objTable.Columns[2].Width = 100;
                objTable.Columns[3].Width = 100;
                objTable.Columns[4].Width = 100;
                objTable.Columns[5].Width = 60;



                objTable.Cell(1, 1).Range.Text = "Date & day";
                objTable.Cell(1, 2).Range.Text = "Publication";
                objTable.Cell(1, 3).Range.Text = "Publication type";
                objTable.Cell(1, 4).Range.Text = "Edition";
                objTable.Cell(1, 5).Range.Text = "Page No. ";

                objTable.Cell(2, 1).Range.Text = objNewsArticle.NewsDate.ToString("dd MMM yyyy, dddd");
                objTable.Cell(2, 2).Range.Text = objNewsArticle.Publication;
                objTable.Cell(2, 3).Range.Text = objNewsArticle.PublicationType;
                objTable.Cell(2, 4).Range.Text = objNewsArticle.Edition;
                objTable.Cell(2, 5).Range.Text = objNewsArticle.Pageno;

                objTable.Cell(3, 1).Range.Text = "Journalist/Source";
                objTable.Cell(3, 2).Range.Text = "Language";
                objTable.Cell(3, 3).Range.Text = "Sentiment";
                objTable.Cell(3, 4).Merge(objTable.Cell(3, 5));
                objTable.Cell(3, 4).Range.Text = "Article type";

                objTable.Cell(4, 1).Range.Text = objNewsArticle.Journalist;
                objTable.Cell(4, 2).Range.Text = objNewsArticle.Language;
                objTable.Cell(4, 3).Range.Text = objNewsArticle.Sentiment;
                objTable.Cell(4, 4).Merge(objTable.Cell(4, 5));
                objTable.Cell(4, 4).Range.Text = objNewsArticle.ArticleType;


                objTable.Cell(5, 1).Range.Text = "Circulation";
                objTable.Cell(5, 2).Merge(objTable.Cell(5, 3));
                objTable.Cell(5, 2).Range.Text = "AVE";

                objTable.Cell(5, 3).Merge(objTable.Cell(5, 4));
                objTable.Cell(5, 3).Range.Text = "Impact score";


                objTable.Cell(6, 1).Range.Text = objNewsArticle.CirculationScore;

                objTable.Cell(6, 2).Merge(objTable.Cell(6, 3));
                objTable.Cell(6, 2).Range.Text = objNewsArticle.AVE;
                objTable.Cell(6, 3).Merge(objTable.Cell(6, 4));
                objTable.Cell(6, 3).Range.Text = objNewsArticle.pi_score;







                objTable.Cell(7, 1).Range.Text = "Headline: ";
                //objTable.Cell(3, 2).Merge(objTable.Cell(3, 3));
                //objTable.Cell(3, 3).Merge(objTable.Cell(3, 4));
                objTable.Cell(7, 2).Merge(objTable.Cell(7, 5));
                objTable.Cell(7, 2).Range.Text = objNewsArticle.HeadLine;
                //objTable.Cell(3, 4).Merge(objTable.Cell(3, 5));
                Object oRange = objTable.Cell(7, 2).Range;

                object linkAddrNewsURL = objNewsArticle.NewsURL;
                object HeadlineNewsURL = objNewsArticle.HeadLine;
                wrdRng.Hyperlinks.Add(oRange, ref linkAddrNewsURL, ref oMissing, ref HeadlineNewsURL, ref HeadlineNewsURL, ref oMissing);

                objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                objTable.Rows[3].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                objTable.Rows[5].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                objTable.Rows[1].Range.Font.Bold = 1;
                objTable.Rows[3].Range.Font.Bold = 1;
                objTable.Rows[5].Range.Font.Bold = 1;
                objTable.Cell(7, 1).Range.Font.Bold = 1;
                objTable.Cell(7, 1).Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);



                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.NewsDate)))
                //{
                //    TextBold(objTable, objNewsArticle.NewsDate.ToString("dd MMMM yyyy, dddd"), 1, 1);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Publication)))
                //{
                //    TextBold(objTable, objNewsArticle.Publication, 1, 2);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.PublicationType)))
                //{
                //    TextBold(objTable, objNewsArticle.PublicationType, 1, 3);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Edition)))
                //{
                //    TextBold(objTable, objNewsArticle.Edition, 1, 4);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Pageno)))
                //{
                //    TextBold(objTable, objNewsArticle.Pageno, 1, 5);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Journalist)))
                //{
                //    TextBold(objTable, objNewsArticle.Journalist, 2, 1);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Language)))
                //{
                //    TextBold(objTable, objNewsArticle.Language, 2, 2);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Sentiment)))
                //{
                //    TextBold(objTable, objNewsArticle.Sentiment, 2, 3);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.ArticleType)))
                //{
                //    TextBold(objTable, objNewsArticle.ArticleType, 2, 4);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.HeadLine)))
                //{
                //    objTable.Cell(3, 2).Range.Font.Bold = 1;
                //}
                //// Set cell shading (background color)
                //objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                //objTable.Rows[2].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Colomncolor);

                // Enable borders for cells and columns
                objTable.Range.Rows.Borders.Enable = 1;
                objTable.Range.Columns.Borders.Enable = 1;
                objTable.Columns.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                objTable.Columns.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                // Set border colors
                objTable.Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                objTable.Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                //// Format the first row (header)
                //objTable.Rows[1].Range.Font.Bold = 1;
                //objTable.Rows[1].Range.Font.ColorIndex = WdColorIndex.wdBlack;

                // Add fine border lines to specific cells (Category, Topic, and Sentiment)
                objTable.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;

                objTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                objTable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            }
            catch (Exception)
            {

                throw;
            }
        }

        public void InsertHeaderTable_New_Online(NewsArticle_Online objNewsArticle)
        {

            try
            {

                int i = 0;
                int j = 0;
                Microsoft.Office.Interop.Word.Table objTable;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                objTable = oDoc.Tables.Add(wrdRng, 5, 5, ref oMissing, ref oMissing);
                objTable.Range.ParagraphFormat.SpaceAfter = 1;
                objTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                // Set column widths (adjust as needed)
                objTable.Columns[1].Width = 100;
                objTable.Columns[2].Width = 100;
                objTable.Columns[3].Width = 100;
                objTable.Columns[4].Width = 100;
                objTable.Columns[5].Width = 60;

                objTable.Cell(1, 1).Range.Text = "Date & day";
                objTable.Cell(1, 2).Range.Text = "News website";
                objTable.Cell(1, 3).Range.Text = "Publication type";
                objTable.Cell(1, 4).Range.Text = "Domain Authority";
                objTable.Cell(1, 5).Range.Text = "Site Traffic";

                objTable.Cell(2, 1).Range.Text = objNewsArticle.NewsDate.ToString("dd MMM yyyy, dddd");
                objTable.Cell(2, 2).Range.Text = objNewsArticle.Publication;
                objTable.Cell(2, 3).Range.Text = objNewsArticle.PublicationType;
                objTable.Cell(2, 4).Range.Text = objNewsArticle.DA;
                if (objNewsArticle.NewsTraffic == null)
                {
                    objTable.Cell(2, 5).Range.Text = "NA";

                }
                else
                {
                    objTable.Cell(2, 5).Range.Text = objNewsArticle.NewsTraffic;
                }



                objTable.Cell(3, 1).Range.Text = "Journalist/Source";
                objTable.Cell(3, 2).Range.Text = "Language";
                objTable.Cell(3, 3).Range.Text = "Sentiment";
                //objTable.Cell(3, 4).Merge(objTable.Cell(3, 5));

                objTable.Cell(3, 4).Range.Text = "Article type";
                objTable.Cell(3, 5).Range.Text = "Impact score";

                objTable.Cell(4, 1).Range.Text = objNewsArticle.Journalist;
                objTable.Cell(4, 2).Range.Text = objNewsArticle.Language;
                objTable.Cell(4, 3).Range.Text = objNewsArticle.Sentiment;
                // objTable.Cell(4, 4).Merge(objTable.Cell(4, 5));
                objTable.Cell(4, 5).Range.Text = objNewsArticle.pi_score;
                if (objNewsArticle.ArticleType.EndsWith(", "))
                {
                    string ArticleType;
                    //objNewsArticle.ArticleType
                    ArticleType = objNewsArticle.ArticleType.Substring(0, objNewsArticle.ArticleType.Length - 2);
                    objTable.Cell(4, 4).Range.Text = ArticleType; //objNewsArticle.ArticleType;
                }


                objTable.Cell(5, 1).Range.Text = "Headline: ";
                //objTable.Cell(3, 2).Merge(objTable.Cell(3, 3));
                //objTable.Cell(3, 3).Merge(objTable.Cell(3, 4));
                objTable.Cell(5, 2).Merge(objTable.Cell(5, 5));
                objTable.Cell(5, 2).Range.Text = objNewsArticle.NewsURL;


                object linkAddrNewsURL = objNewsArticle.NewsURL;
                object HeadlineNewsURL = objNewsArticle.Title;

                //  Microsoft.Office.Interop.Word.Selection wrdSelection1 = oWord.ActiveWindow.Selection;
                Object oRange = objTable.Cell(5, 2).Range;
                wrdRng.Hyperlinks.Add(oRange, ref linkAddrNewsURL, ref oMissing, ref HeadlineNewsURL, ref HeadlineNewsURL, ref oMissing);

                objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                objTable.Rows[3].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                objTable.Rows[1].Range.Font.Bold = 1;
                objTable.Rows[3].Range.Font.Bold = 1;
                objTable.Cell(5, 1).Range.Font.Bold = 1;
                objTable.Cell(5, 1).Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);

                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.NewsDate)))
                //{
                //    TextBold(objTable, objNewsArticle.NewsDate.ToString("dd MMMM yyyy, dddd"), 1, 1);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Publication)))
                //{
                //    TextBold(objTable, objNewsArticle.Publication, 1, 2);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.PublicationType)))
                //{
                //    TextBold(objTable, objNewsArticle.PublicationType, 1, 3);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.DA)))
                //{
                //    TextBold(objTable, objNewsArticle.DA, 1, 4);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.NewsTraffic)))
                //{
                //    TextBold(objTable, objNewsArticle.NewsTraffic, 1, 5);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Journalist)))
                //{
                //    TextBold(objTable, objNewsArticle.Journalist, 2, 1);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Language)))
                //{
                //    TextBold(objTable, objNewsArticle.Language, 2, 2);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.Sentiment)))
                //{
                //    TextBold(objTable, objNewsArticle.Sentiment, 2, 3);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.ArticleType)))
                //{
                //    TextBold(objTable, objNewsArticle.ArticleType, 2, 4);
                //}
                //if (!string.IsNullOrEmpty(Convert.ToString(objNewsArticle.NewsURL)))
                //{
                //    objTable.Cell(3, 2).Range.Font.Bold = 1;
                //}


                // Enable borders for cells and columns
                // Enable borders for cells and columns
                objTable.Range.Rows.Borders.Enable = 1;
                objTable.Range.Columns.Borders.Enable = 1;
                objTable.Columns.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                objTable.Columns.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                // Set border colors
                objTable.Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                objTable.Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                //// Format the first row (header)
                //objTable.Rows[1].Range.Font.Bold = 1;
                //objTable.Rows[1].Range.Font.ColorIndex = WdColorIndex.wdBlack;

                // Add fine border lines to specific cells (Category, Topic, and Sentiment)
                objTable.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;

                objTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                objTable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            }
            catch (Exception)
            {

                throw;
            }

        }

        public void InsertTable_Articles_Print(List<NewsArticle> lstNewsArticle)
        {
            try
            {

                int i = 0;
                int j = 0;
                int rowCount = lstNewsArticle.Count + 1;
                int columnCount = 5;

                Microsoft.Office.Interop.Word.Table objTable;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                objTable = oDoc.Tables.Add(wrdRng, rowCount, columnCount, ref oMissing, ref oMissing);
                objTable.Range.ParagraphFormat.SpaceAfter = 1;
                objTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                objTable.Range.Font.Size = 12.0f;


                // Set column widths (adjust as needed)
                objTable.Columns[1].Width = 100;
                objTable.Columns[2].Width = 180;
                objTable.Columns[3].Width = 80;
                objTable.Columns[4].Width = 80;
                objTable.Columns[5].Width = 40;

                // Uncomment and adjust if you want columns 6 to 9
                // objTable.Columns[6].Width = 120;
                // objTable.Columns[7].Width = 120;
                // objTable.Columns[8].Width = 120;
                // objTable.Columns[9].Width = 80;

                // Uncomment if you want a header row
                // Add heading row
                // objTable.Rows[1].Range.Text = "News Table";
                // objTable.Rows[1].Range.Cells.Merge();
                // objTable.Rows[1].Range.Font.Bold = 1;
                // objTable.Rows[1].Range.Font.ColorIndex = WdColorIndex.wdWhite;
                objTable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // Add columns headers
                objTable.Cell(1, 1).Range.Text = "Article Date";
                objTable.Cell(1, 2).Range.Text = "Headline";
                objTable.Cell(1, 3).Range.Text = "Publication";
                objTable.Cell(1, 4).Range.Text = "Edition";
                objTable.Cell(1, 5).Range.Text = "Page No.";

                //objTable.Cell(1, 1).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 2).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 3).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 4).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 5).Range.Font.ColorIndex = WdColorIndex.wdWhite;


                objTable.Cell(1, 1).Range.Font.Bold = 1;
                objTable.Cell(1, 2).Range.Font.Bold = 1;
                objTable.Cell(1, 3).Range.Font.Bold = 1;
                objTable.Cell(1, 4).Range.Font.Bold = 1;
                objTable.Cell(1, 5).Range.Font.Bold = 1;


                //Uncomment if you want borders for columns
                //Enable borders for columns
                for (int m = 1; m <= columnCount; m++)
                {
                    objTable.Columns[m].Borders.Enable = 1;
                    objTable.Columns[m].Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                    objTable.Columns[m].Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                }

                Int16 No = 1;

                Microsoft.Office.Interop.Word.Selection wrdSelection = oWord.ActiveWindow.Selection;

                // Add articles to the table
                for (int n = 1; n < rowCount; n++)
                {
                    object linkAddr = "#B_Print" + (No - 1).ToString();
                    object Headline = lstNewsArticle[n - 1].HeadLine;

                    Object oRange = objTable.Cell(No + 1, 2).Range;
                    wrdSelection.Hyperlinks.Add(oRange, ref linkAddr, ref oMissing, ref Headline, ref Headline, ref oMissing);

                    objTable.Cell(No + 1, 1).Range.Text = lstNewsArticle[n - 1].NewsDate.ToString("dd MMM, yyyy");
                   // objTable.Cell(No + 1, 1).Range.Font.Bold = 1;
                    //objTable.Cell(n + 2, 2).Range.Text = lstNewsArticle[n].HeadLine;
                    objTable.Cell(No + 1, 3).Range.Text = lstNewsArticle[n - 1].Publication;
                    objTable.Cell(No + 1, 4).Range.Text = lstNewsArticle[n - 1].Edition;
                    objTable.Cell(No + 1, 5).Range.Text = lstNewsArticle[n - 1].Pageno;
                    objTable.Cell(No + 1, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    objTable.Cell(No + 1, 5).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                     objTable.Cell(n + 2, 5).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    No++;
                }


                // Set vertical alignment for the specified cell


                //objTable.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Color.Blue);

                // Set cell shading (background color)
                objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                //for (int p = 0; p < rowCount; p++)
                //{
                //    objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                //    if (p >= 1)
                //    {
                //        if (p % 2 == 0)
                //        {
                //            objTable.Rows[p].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(LightBlueColor);
                //        }
                //    }
                //}

                // Enable borders for cells and columns
                for (int k = 0; k < rowCount; k++)
                {
                    objTable.Rows[k + 1].Range.Borders.Enable = 1;
                }

                // Set border colors
                objTable.Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                objTable.Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                // Add fine border lines to specific cells
                for (int l = 1; l < columnCount; l++)
                {
                    objTable.Cell(2, l).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    objTable.Cell(2, l).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                    // Align text and set vertical alignment
                    objTable.Columns[l].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        public void InsertTable_Articles_Online(List<NewsArticle_Online> lstNewsArticle)
        {

            try
            {
                int i = 0;
                int j = 0;
                int rowCount = lstNewsArticle.Count + 1;
                int columnCount = 3;

                Microsoft.Office.Interop.Word.Table objTable;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                objTable = oDoc.Tables.Add(wrdRng, rowCount, columnCount, ref oMissing, ref oMissing);
                objTable.Range.ParagraphFormat.SpaceAfter = 1;
                objTable.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                objTable.Range.Font.Size = 12.0f;


                // Set column widths (adjust as needed)
                objTable.Columns[1].Width = 80;
                objTable.Columns[2].Width = 200;
                objTable.Columns[3].Width = 80;
               // objTable.Columns[4].Width = 80;
                //objTable.Columns[5].Width = 40;

                // Uncomment and adjust if you want columns 6 to 9
                // objTable.Columns[6].Width = 120;
                // objTable.Columns[7].Width = 120;
                // objTable.Columns[8].Width = 120;
                // objTable.Columns[9].Width = 80;

                // Uncomment if you want a header row
                // Add heading row
                // objTable.Rows[1].Range.Text = "News Table";
                // objTable.Rows[1].Range.Cells.Merge();
                // objTable.Rows[1].Range.Font.Bold = 1;
                // objTable.Rows[1].Range.Font.ColorIndex = WdColorIndex.wdWhite;
                objTable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // Add columns headers
                objTable.Cell(1, 1).Range.Text = "Article Date";
                objTable.Cell(1, 2).Range.Text = "Headline";
                objTable.Cell(1, 3).Range.Text = "Publication";
               // objTable.Cell(1, 4).Range.Text = "Edition";
                //objTable.Cell(1, 5).Range.Text = "Page No";

                //objTable.Cell(1, 1).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 2).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 3).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 4).Range.Font.ColorIndex = WdColorIndex.wdWhite;
                //objTable.Cell(1, 5).Range.Font.ColorIndex = WdColorIndex.wdWhite;


                objTable.Cell(1, 1).Range.Font.Bold = 1;
                objTable.Cell(1, 2).Range.Font.Bold = 1;
                objTable.Cell(1, 3).Range.Font.Bold = 1;
               // objTable.Cell(1, 4).Range.Font.Bold = 1;
                //objTable.Cell(1, 5).Range.Font.Bold = 1;


                //Uncomment if you want borders for columns
                //Enable borders for columns
                for (int m = 1; m <= columnCount; m++)
                {
                    objTable.Columns[m].Borders.Enable = 1;
                    objTable.Columns[m].Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                    objTable.Columns[m].Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                }


                Int16 No = 1;

                Microsoft.Office.Interop.Word.Selection wrdSelection = oWord.ActiveWindow.Selection;

                // Add articles to the table
                for (int n = 1; n < rowCount; n++)
                {
                    object linkAddr = "#B_Online" + (No - 1).ToString();
                    object Headline = lstNewsArticle[n - 1].Title;

                    Object oRange = objTable.Cell(No + 1, 2).Range;
                    wrdSelection.Hyperlinks.Add(oRange, ref linkAddr, ref oMissing, ref Headline, ref Headline, ref oMissing);

                    objTable.Cell(No + 1, 1).Range.Text = lstNewsArticle[n - 1].NewsDate.ToString("dd MMM, yyyy");
                    //objTable.Cell(n + 2, 2).Range.Text = lstNewsArticle[n].Title;
                    objTable.Cell(No + 1, 3).Range.Text = lstNewsArticle[n - 1].Publication;
                   // objTable.Cell(No + 1, 4).Range.Text = lstNewsArticle[n - 1].Edition;
                    //objTable.Cell(n + 2, 5).Range.Text = lstNewsArticle[n].Pageno;

                    No++;
                }


                //objTable.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Color.Blue);

                // Set cell shading (background color)
                objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                //for (int p = 0; p < rowCount; p++)
                //{
                //    objTable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);

                //    if (p >= 1)
                //    {
                //        if (p % 2 == 0)
                //        {
                //            objTable.Rows[p].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(LightBlueColor);
                //        }
                //    }
                //}

                // Enable borders for cells and columns
                for (int k = 0; k < rowCount; k++)
                {
                    objTable.Rows[k + 1].Range.Borders.Enable = 1;
                }

                // Set border colors
                objTable.Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                objTable.Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                // Add fine border lines to specific cells
                for (int l = 1; l < columnCount; l++)
                {
                    objTable.Cell(2, l).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    objTable.Cell(2, l).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                    // Align text and set vertical alignment
                    objTable.Columns[l].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

            }
            catch (Exception)
            {

                throw;
            }
        }


        public void GotoastLinePoint()
        {
            object what = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPercent;
            object which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToLast;
            oWord.Selection.GoTo(ref what, ref which, oMissing, oMissing);
        }

        public void GotoFirstLinePoint()
        {
            object what = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine;
            object which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst;
            oWord.Selection.GoTo(ref what, ref which, oMissing, oMissing);
        }

        public void InsertDomainImageIntoBody(string URL)
        {
            if (!string.IsNullOrEmpty(URL))
            {
                try
                {
                    oDoc.Shapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing).Left = (float)WdFramePosition.wdFrameCenter;
                    Microsoft.Office.Interop.Word.InlineShape oShape;
                    object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oShape = wrdRng.InlineShapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing);
                    //oShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    //oShape.Width = (float)250;

                    oShape.Width = 20; // Replace with the desired width in points
                    oShape.Height = 20; // Replace with the desired height in points

                    oWord.ActiveWindow.Selection.InsertAfter("\n");
                    oWord.ActiveWindow.Selection.InsertAfter("\n");

                }
                catch (Exception ex)
                {

                }
            }
            //oWord.ActiveWindow.Selection.InsertAfter("\n");

        }
        public void InsertDomainImageIntoBody_New(string imagePath)
        {
            if (!string.IsNullOrEmpty(imagePath))
            {
                try
                {

                    //  oWord.Selection.InlineShapes.AddPicture(imagePath, oMissing, oMissing, oMissing);
                    //wrdRng.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                    //wrdRng.Borders.OutsideLineWidth = (WdLineWidth)WdLineWidth.wdLineWidth100pt;
                    word.InlineShape inlineShape = oWord.Selection.InlineShapes.AddPicture(imagePath, oMissing, oMissing, oMissing);

                    // oShape = wrdRng.InlineShapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing);
                    inlineShape.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    inlineShape.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth100pt;
                    inlineShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


                    float maxHeight = 270.0f;

                    // Check if the current height exceeds the maximum height
                    if (inlineShape.Height > maxHeight)
                    {
                        // Calculate the ratio to maintain the original aspect ratio
                        //   float ratio = maxHeight / inlineShape.Height;

                        // Adjust the height and width while maintaining the aspect ratio
                        inlineShape.Height = maxHeight;
                        //   inlineShape.Width *= ratio;
                    }

                    oWord.ActiveWindow.Selection.InsertAfter("\n");
                }
                catch (Exception ex)
                {

                }
            }

        }
        public void InsertDomainImageIntoBody_New_4Graph(string imagePath)
        {
            if (!string.IsNullOrEmpty(imagePath))
            {
                try
                {

                    //  oWord.Selection.InlineShapes.AddPicture(imagePath, oMissing, oMissing, oMissing);
                    //wrdRng.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                    //wrdRng.Borders.OutsideLineWidth = (WdLineWidth)WdLineWidth.wdLineWidth100pt;
                    word.InlineShape inlineShape = oWord.Selection.InlineShapes.AddPicture(imagePath, oMissing, oMissing, oMissing);

                    // oShape = wrdRng.InlineShapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing);
                    inlineShape.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    inlineShape.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth100pt;
                    inlineShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


                    float maxHeight = 270.0f;

                    // Check if the current height exceeds the maximum height
                    if (inlineShape.Height > maxHeight)
                    {
                        // Calculate the ratio to maintain the original aspect ratio
                        //   float ratio = maxHeight / inlineShape.Height;

                        // Adjust the height and width while maintaining the aspect ratio
                        inlineShape.Height = maxHeight;
                        //   inlineShape.Width *= ratio;
                    }
                    inlineShape.Width = 320.0f; 
                    oWord.ActiveWindow.Selection.InsertAfter("\n");
                }
                catch (Exception ex)
                {

                }
            }

        }


        public void ReplcaeBookMarkWithText(string Bookmark, string Text)
        {
            try
            {
                //oDoc.Bookmarks[Bookmark].Select();
                //oDoc.Bookmarks[Bookmark].Range.Text = Text;
                Microsoft.Office.Interop.Word.Paragraph para = oDoc.Paragraphs.Add(ref oMissing);
                object oRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
                wrdRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
                wrdRng.Text = Text;
                try
                {
                    //BOOKMARKING
                    if (!string.IsNullOrEmpty(Bookmark))
                    {
                        oDoc.Bookmarks.Add(Bookmark, oMissing);
                    }
                }
                catch (Exception ex)
                {

                }
            }
            catch (Exception ex) { throw ex; }
        }
        public void InertText(string text, string Bookmark)
        {
            try
            {
                //oWord.Documents.Add(@"D:\Dosier\WordDocWithTemplate\WordDocWithTemplate\Templates\Template.dotx");
                //object oRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
                //oDoc.Bookmarks["MyText"].Select();
                //oDoc.Bookmarks["MyText"].Range.Text = text;
                //oWord.Selection.TypeText(text);
                //Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
                //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                //wrdRng.Text = text;
            }
            catch (Exception ex) { throw ex; }
        }
        public void InsertArticleImageIntoBody(string URL, string Bookmark)
        {
            //object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng.            

            //oDoc.Shapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing).Left = (float)WdFramePosition.wdFrameCenter;

            //oWord.ActiveWindow.Selection.InsertAfter("\n");

            Microsoft.Office.Interop.Word.InlineShape oShape;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            try
            {
                URL = URL.Replace("\r\n", "");
                //  URL = "http://mtimages.merryspiders.com/ArticleImages/Uploaded_Images/2024/1/18/2024_1_18_2b8511b7-e994-449d-a87f-2f12a24bfcc7.jpg";


                word.InlineShape inlineShape = oWord.Selection.InlineShapes.AddPicture(URL, oMissing, oMissing, oMissing);

                // oShape = wrdRng.InlineShapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing);
                inlineShape.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                inlineShape.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth100pt;
                // Center-align the image within the paragraph
                inlineShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                float maxHeight = 340.0f;

                // Check if the current height exceeds the maximum height
                if (inlineShape.Height > maxHeight)
                {
                    // Calculate the ratio to maintain the original aspect ratio
                    // float ratio = maxHeight / inlineShape.Height;

                    // Adjust the height and width while maintaining the aspect ratio
                    inlineShape.Height = maxHeight;
                    // inlineShape.Width *= ratio;
                }

                // wrdRng.Borders.OutsideLineWidth = (WdLineWidth)WdLineWidth.wdLineWidth100pt;
            }
            catch (Exception ex)
            {
                word.InlineShape inlineShape = oWord.Selection.InlineShapes.AddPicture(ConfigurationManager.AppSettings["Imagenotavailable"], oMissing, oMissing, oMissing);




                inlineShape.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                inlineShape.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth100pt;
                // oShape = wrdRng.InlineShapes.AddPicture(ConfigurationManager.AppSettings["Imagenotavailable"], ref oMissing, ref oMissing, ref oMissing);
            }
            //  oShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //oShape.Height = auto

            try
            {
                //BOOKMARKING
                if (!string.IsNullOrEmpty(Bookmark))
                {
                    oDoc.Bookmarks.Add(Bookmark, oMissing);
                }
            }
            catch (Exception ex)
            {

            }
            //oWord.ActiveWindow.Selection.InsertAfter("\n");
        }




        //public void InsertNewsArticleImageIntoBody(string URL, string Bookmark)
        //{
        //    Microsoft.Office.Interop.Word.InlineShape oShape;
        //    object oRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
        //    Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
        //    wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
        //    try
        //    {
        //        oShape = wrdRng.InlineShapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing);
        //    }
        //    catch (Exception ex)
        //    {
        //        oShape = wrdRng.InlineShapes.AddPicture(ConfigurationManager.AppSettings["Imagenotavailable"], ref oMissing, ref oMissing, ref oMissing);
        //    }
        //    oShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //    //oShape.Height = auto

        //    try
        //    {
        //        //BOOKMARKING
        //        if (!string.IsNullOrEmpty(Bookmark))
        //        {
        //            oDoc.Bookmarks.Add(Bookmark, oMissing);
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    //oWord.ActiveWindow.Selection.InsertAfter("\n");
        //}

        public void InsertNewsArticleImageIntoBody(string Bookmark, string URL)
        {
            //object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //wrdRng.            

            //oDoc.Shapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing).Left = (float)WdFramePosition.wdFrameCenter;

            //oWord.ActiveWindow.Selection.InsertAfter("\n");

            Microsoft.Office.Interop.Word.InlineShape oShape;
            object oRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(Bookmark).Range;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            try
            {
                oShape = wrdRng.InlineShapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing);
            }
            catch (Exception ex)
            {
                oShape = wrdRng.InlineShapes.AddPicture(ConfigurationManager.AppSettings["Imagenotavailable"], ref oMissing, ref oMissing, ref oMissing);
            }
            //oShape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //oShape.Height = auto

            try
            {
                //BOOKMARKING
                if (!string.IsNullOrEmpty(Bookmark))
                {
                    oDoc.Bookmarks.Add(Bookmark, oMissing);
                }
            }
            catch (Exception ex)
            {

            }
            //oWord.ActiveWindow.Selection.InsertAfter("\n");
        }

        public void InsertBoldText(string strBookMark, string Text)
        {
            //oWord.Documents.Add(@"D:\Dosier\WordDocWithTemplate\WordDocWithTemplate\Templates\Template.dotx");
            //object oRng = oDoc.Bookmarks.get_Item(strBookMark).Range;
            //oDoc.Bookmarks[strBookMark].Select();
            //oDoc.Bookmarks[strBookMark].Range.Text = Text;
            //oWord.Selection.TypeText(Text);
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(strBookMark).Range;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            wrdRng.Text = Text;
            wrdRng.Select();
            wrdRng.Font.Bold = 1;

            if (!string.IsNullOrEmpty(strBookMark))
            {
                oDoc.Bookmarks.Add(strBookMark, oMissing);
            }
        }
        public void InsertBoldTextLable(string lable)
        {
     
            oWord.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;

            oWord.Selection.Font.Bold = 2;
            oWord.Selection.TypeText(lable);

           
        }

        public void InsertBoldText(string strBookMark, string Label, string Text)
        {
            try
            {
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(strBookMark).Range;
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.Text = Label + Text;
                //wrdRng.Text = Text;


                object objStart = wrdRng.Start;
                object objEnd = wrdRng.Start + Label.IndexOf(":");

                Microsoft.Office.Interop.Word.Range rngBold = oDoc.Range(ref objStart, ref objEnd);
                rngBold.Bold = 1;

                if (!string.IsNullOrEmpty(strBookMark))
                {
                    oDoc.Bookmarks.Add(strBookMark, oMissing);
                }
            }
            catch (Exception ex) { throw ex; }
        }

        public void InsertText(string strBookMark, string Text)
        {
            //oWord.Documents.Add(@"D:\Dosier\WordDocWithTemplate\WordDocWithTemplate\Templates\Template.dotx");
            //object oRng = oDoc.Bookmarks.get_Item(strBookMark).Range;
            //oDoc.Bookmarks[strBookMark].Select();
            //oDoc.Bookmarks[strBookMark].Range.Text = Text;
            //oWord.Selection.TypeText(Text);
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(strBookMark).Range;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.Text = Text;
            //wrdRng.Font.Bold = 1;

            if (!string.IsNullOrEmpty(strBookMark))
            {
                oDoc.Bookmarks.Add(strBookMark, oMissing);
            }



        }
        public void CreateIndexPage()
        {

        }
        public void SaveAs(string strFileName)
        {
            object fileName = strFileName;

            //oDoc.SaveAs(ref fileName, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            oDoc.Application.ActiveDocument.SaveAs(ref fileName, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }

        public void Quit()
        {
            oDoc.Close();
            oWord.Application.Quit(ref oMissing, ref oMissing, ref oMissing);
        }

        public void ReplaceText(string strReplace, string strToReplace)
        {
            try
            {
                //var app = new Application();
                //var doc = app.Documents.Add(
                //doc.Activate();
                //var keyword = "logo-o";
                var sel = oWord.Selection;
                sel.Find.Text = string.Format("[{0}]", strReplace);
                sel.Find.Replacement.Text = strReplace;
                sel.Find.Execute(Replace: WdReplace.wdReplaceOne);
                //sel.Find.Execute(Replace: WdReplace.wdReplaceNone);
                //sel.Range.Select();
                //var imgPath = string.Format(@"D:\Bootstrap POC\products-WB0573SK0\SmartAdmin_1_3\DEVELOPER\HTML_version\img\{0}", strReplace);
                //sel.InlineShapes.AddPicture(FileName: imgPath, LinkToFile: false, SaveWithDocument: true);
                //wordAutomation 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void ReplaceTextWithBreak(string strReplace)
        {
            try
            {
                //var app = new Application();
                //var doc = app.Documents.Add(
                //doc.Activate();
                //var keyword = "logo-o";
                var sel = oWord.Selection;
                sel.Find.Text = string.Format("{0}", strReplace);
                sel.Find.Replacement.Text = "^l3";
                sel.Find.Execute(Replace: WdReplace.wdReplaceAll);
                //sel.Find.Execute(Replace: WdReplace.wdReplaceNone);
                //sel.Range.Select();
                //var imgPath = string.Format(@"D:\Bootstrap POC\products-WB0573SK0\SmartAdmin_1_3\DEVELOPER\HTML_version\img\{0}", strReplace);
                //sel.InlineShapes.AddPicture(FileName: imgPath, LinkToFile: false, SaveWithDocument: true);
                //wordAutomation 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void PageHeading(string Text, string objBookmark)
        {
            // Insert a horizontal line at the end of the document with color
            Microsoft.Office.Interop.Word.Paragraph paragraph = oDoc.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Borders borders = paragraph.Borders;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            try
            {
                //BOOKMARKING
                if (!string.IsNullOrEmpty(objBookmark))
                {
                    oDoc.Bookmarks.Add(objBookmark, oRng);

                }
            }
            catch (Exception ex)
            {

            }

            paragraph.Range.InsertParagraphAfter();

            paragraph.Range.Text = Text;
            // Explicitly set this to "not bold"
            paragraph.Range.Font.Bold = 7;
            paragraph.Range.Font.Size = 15;


            // Set line style to single, color to blue, and line width (adjust as needed)
            borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            borders[WdBorderType.wdBorderBottom].Color = (WdColor)ColorTranslator.ToOle(Orangecolor);
            borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt; // Adjust line width

            // Move to the next line for content
            oWord.Selection.TypeParagraph();


        }
        public void PageHeading_New(string Text, string objBookmark)
        {


            Microsoft.Office.Interop.Word.Table table1;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table1 = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);
            table1.Range.ParagraphFormat.SpaceAfter = 1;

            //BOOKMARKING
            if (!string.IsNullOrEmpty(objBookmark))
            {
                oDoc.Bookmarks.Add(objBookmark, wrdRng);
            }

            table1.Cell(1, 1).Range.Text = Text;
            table1.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            // Set vertical alignment for the specified cell
            table1.Cell(1, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;


            table1.Cell(1, 1).Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
            table1.Cell(1, 1).Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);

            table1.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
            table1.Rows[1].Range.Font.Bold = 600;
            table1.Cell(1, 1).Range.Font.Size = 15.0f;

            table1.Cell(1, 1).Height = 22;
            table1.Cell(1, 1).Height = 22;
            GotoastLinePoint();




        }

        public void FirstPage_New(string clientName, string Description, string Dates, string imagePath)
        {

            if (!string.IsNullOrEmpty(imagePath))
            {
                if (!string.IsNullOrEmpty(imagePath))
                {
                    try
                    {

                        // Add picture to the document
                        word.InlineShape inlineShape = oWord.Selection.InlineShapes.AddPicture(imagePath, oMissing, oMissing, oMissing);
                        // float desiredWidth = 420.0f; // set your desired width in points
                        float desiredHeight = 650.0f; // set your desired height in points

                        // inlineShape.Width = desiredWidth;
                        inlineShape.Height = desiredHeight;
                        // Set the image as a background image
                        inlineShape.Select();
                        oWord.Selection.ShapeRange.WrapFormat.Type = word.WdWrapType.wdWrapBehind;

                        GotoastLinePoint();
                        InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();

                        GotoastLinePoint();
                        InserLineBreak();




                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();
                        //GotoastLinePoint();
                        //InserLineBreak();

                        GotoastLinePoint();
                        InserLineBreak();

                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        GotoastLinePoint();
                        InserLineBreak();
                        // Set the selection range to the end of the document
                        word.Range range = oWord.ActiveDocument.Range(oWord.Selection.End, oWord.Selection.End);

                        range.InsertParagraphAfter();



                        range.Text = "   " + clientName;
                        range.Font.Bold = 7;
                        range.Font.Size = 28;
                        range.Font.Color = WdColor.wdColorWhite;
                        // Explicitly set this to "not bold"
                        GotoastLinePoint();
                        InserLineBreak();

                        word.Range rangeDescription = oWord.ActiveDocument.Range(oWord.Selection.End, oWord.Selection.End);

                       
                        
                        // Explicitly set this to "not bold"
                        rangeDescription.Text = "   Topic: " + Description;
                        rangeDescription.Font.Bold = 3;
                        rangeDescription.Font.Size = 16;
                       
                      //  rangeDescription.ParagraphFormat.RightIndent = oWord.InchesToPoints(0.5f); 
                       // rangeDescription.ParagraphFormat.Alignment = word.WdParagraphAlignment.wdAlignParagraphJustify;
                        rangeDescription.Font.Color = WdColor.wdColorWhite;
                        GotoastLinePoint();
                        InserLineBreak();



                        
                       

                        word.Range rangeDates = oWord.ActiveDocument.Range(oWord.Selection.End, oWord.Selection.End);
                        // Explicitly set this to "not bold"
                        rangeDates.Font.Bold = 3;
                        rangeDates.Font.Size = 18;
                        rangeDates.Font.Color = WdColor.wdColorWhite;
                        
                        rangeDates.Text = "   " + Dates;
                     //   rangeDates.ParagraphFormat.RightIndent = 0;
                        // Insert text after the image
                        //range.Text = "Your Text Here";
                        GotoastLinePoint();
                        InserLineBreak();
                        // Move the selection back to the end of the document
                        oWord.Selection.EndKey(word.WdUnits.wdStory);




                    }
                    catch (Exception ex)
                    {
                        // Handle the exception
                    }
                }
            }


        }
        public void Insert_NewLine()
        {
            oWord.ActiveWindow.Selection.InsertAfter("\n");
        }

        public void InsertVerticleLine()
        {
            Microsoft.Office.Interop.Word.Section section = oDoc.Sections[1]; // Get the first section

            // Add a vertical line in the body of the document
            Microsoft.Office.Interop.Word.Shape verticalLine = oDoc.Shapes.AddLine(
                section.PageSetup.PageWidth / 2, // X1 - Center of the page
                20, // Y1 - Top of the page (adjust as needed)
                section.PageSetup.PageWidth / 2, // X2 - Center of the page
                20 + 2 * 28.35f // Y2 - Top of the page plus 2 cm (28.35 points per cm)
            );
            // Set the line properties
            verticalLine.Line.ForeColor.RGB = (int)WdColor.wdColorBlack; // Adjust the line color as needed
            verticalLine.Line.Weight = 10;
        }

        public void FirstPage(string clientName, string Description, string Dates)
        {
            GotoastLinePoint();
            InsertDocumentHeadline(clientName);
            InserLineBreak();

            GotoastLinePoint();
            InsertText(Description);
            GotoastLinePoint();
            InserLineBreak();
            InserLineBreak();

            GotoastLinePoint();
            InsertText(Dates);


        }

        public void AddImageWithTextToFirstPage(string mainContentImageUrl, string mainContentText)
        {
            foreach (Section section in oDoc.Sections)
            {
                // Check if it is the first page
                if (section.Index == 1)
                {
                    // Get the main content range of the first page
                    Range mainContentRange = section.Range;

                    // Add the main content image
                    InlineShapes inlineShapes = mainContentRange.InlineShapes;
                    InlineShape mainContentImage = inlineShapes.AddPicture(mainContentImageUrl, LinkToFile: false, SaveWithDocument: true);

                    // Set the font properties for the text
                    mainContentRange.Font.Size = 18;
                    mainContentRange.Font.Bold = 1; // Bold
                    mainContentRange.Font.Color = WdColor.wdColorWhite; // White color

                    // Type the text on the main content image
                    mainContentRange.Text = mainContentText;
                }
            }
        }

        //public void Overview_Page()
        //{

        //    CreateSimpleTable(oDoc, 2, 3);

        //    oDoc.Paragraphs.Add();

        //    // Add colspan-like effect in the 4th row
        //    AddColspanEffect(oDoc,2);


        //}

        public void CreateSimpleTable(int rows, int columns, DataTable dt)
        {

            Microsoft.Office.Interop.Word.Table table;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table = oDoc.Tables.Add(wrdRng, rows, columns, ref oMissing, ref oMissing);
            table.Range.ParagraphFormat.SpaceAfter = 1;

            // Set the text in the cells
            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= columns; col++)
                {
                    string Text = string.Empty;
                    if (col == 1)
                    {
                        Text = "ObjectName";
                    }
                    else if (col == 2)
                    {
                        Text = "PrintCount";
                    }
                    else if (col == 3)
                    {
                        Text = "OnlineCount";
                    }

                    table.Cell(row, col).Range.Text = Convert.ToString(dt.Rows[row - 1][Text]);

                    // Center align text in the cell
                    table.Cell(row, col).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }

            // Add borders to the entire table
            table.Borders.Enable = 1; // Enable borders
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle; // Set outside border style
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle; // Set inside border style
        }

        public void AddColspanEffect(int rows, int cols, string text)
        {

            Microsoft.Office.Interop.Word.Table table;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            table = oDoc.Tables.Add(wrdRng, rows, cols, ref oMissing, ref oMissing);
            table.Range.ParagraphFormat.SpaceAfter = 1;

            // Merge cells in the 4th row (2nd index) to simulate colspan
            Row row = table.Rows[rows];
            row.Cells.Merge();
            row.Range.Text = text;
            // Center align text in the cell
            row.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            table.Borders.Enable = 1; // Enable borders
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle; // Set outside border style
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle; // Set inside border style
        }

        public void GotoEndOfDocument()
        {
            oWord.Selection.EndKey(WdUnits.wdStory);
        }

        public void NextLine()
        {
            oWord.ActiveWindow.Selection.InsertAfter("\n");
        }

        public void CreateSimpleTable_1(int rows, int columns, DataTable dt, string objBookmark)
        {
            try
            {

                Microsoft.Office.Interop.Word.Table table;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                table = oDoc.Tables.Add(wrdRng, rows, columns, ref oMissing, ref oMissing);
                table.Range.ParagraphFormat.SpaceAfter = 1;

                try
                {
                    //BOOKMARKING
                    if (!string.IsNullOrEmpty(objBookmark))
                    {
                        oDoc.Bookmarks.Add(objBookmark, wrdRng);
                    }
                }
                catch (Exception ex)
                {

                }

                #region Table
                // Set the text in the cells
                for (int row = 1; row <= rows; row++)
                {
                    GotoastLinePoint();
                    string prev_Sortby = string.Empty;
                    string pres_Sortby = string.Empty;
                    if (row > 1)
                    {
                        prev_Sortby = Convert.ToString(dt.Rows[row - 2]["Sortby"]);
                    }

                    pres_Sortby = Convert.ToString(dt.Rows[row - 1]["Sortby"]);

                    if (row == 1)
                    {
                        //GotoastLinePoint();
                        //table.Rows[row].Cells.Merge();
                        //table.Rows[row].Range.Text = pres_Sortby;
                        //NextLine();
                        //table.Rows[row].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        GotoastLinePoint();
                        for (int col = 1; col <= columns; col++)
                        {
                            string Text = string.Empty;
                            if (col == 1)
                            {
                                Text = "ObjectName";
                            }
                            else if (col == 2)
                            {
                                Text = "PrintCount";
                            }
                            else if (col == 3)
                            {
                                Text = "OnlineCount";
                            }

                            GotoastLinePoint();

                            string originalText = Convert.ToString(dt.Rows[row - 1][Text]);

                            // Check if the string starts with "zTotal" (case-insensitive)
                            if (originalText.StartsWith("Headline", StringComparison.OrdinalIgnoreCase))
                            {
                                // Replace "zTotal" with "Total" at the beginning of the string
                                originalText = originalText.Replace("Headline", "");
                            }

                            table.Cell(row, col).Range.Text = originalText;
                            GotoastLinePoint();
                            // Center align text in the cell
                            table.Cell(row, col).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                            // Set vertical alignment for the specified cell
                            table.Cell(row, col).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                            table.Cell(row, col).Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                            table.Cell(row, col).Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                            table.Cell(row, col).Range.Bold = 1;
                            table.Cell(row, col).Range.Font.Size = 14.0f;

                            table.Cell(row, col).Height = 22;
                            table.Cell(row, col).Height = 22;
                            //table.Cell(row, col).Width = 6.0f;
                        }
                    }
                    else
                    {
                        if (prev_Sortby == pres_Sortby)
                        {
                            for (int col = 1; col <= columns; col++)
                            {
                                string Text = string.Empty;
                                if (col == 1)
                                {
                                    Text = "ObjectName";
                                }
                                else if (col == 2)
                                {
                                    Text = "PrintCount";
                                }
                                else if (col == 3)
                                {
                                    Text = "OnlineCount";
                                }

                                GotoastLinePoint();

                                string originalText = Convert.ToString(dt.Rows[row - 1][Text]);

                                // Check if the string starts with "zTotal" (case-insensitive)
                                if (originalText.StartsWith("zTotal", StringComparison.OrdinalIgnoreCase)
                                    || originalText.StartsWith("zNo", StringComparison.OrdinalIgnoreCase)
                                    || originalText.StartsWith("zzTotal", StringComparison.OrdinalIgnoreCase)
                                    || originalText.StartsWith("zzzCoverage", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Replace "zTotal" with "Total" at the beginning of the string
                                    originalText = originalText.Replace("z", "");
                                    originalText = originalText.Replace("zz", "");
                                    originalText = originalText.Replace("zzz", "");
                                }

                                table.Cell(row, col).Range.Text = originalText;
                                GotoastLinePoint();
                                // Center align text in the cell
                                table.Cell(row, col).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                                // Set vertical alignment for the specified cell
                                table.Cell(row, col).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;


                                table.Cell(row, col).Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                                table.Cell(row, col).Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);

                                if (Convert.ToString(dt.Rows[row - 1][Text]).StartsWith("zTotal", StringComparison.OrdinalIgnoreCase)
                                    || Convert.ToString(dt.Rows[row - 1][Text]).StartsWith("zNo", StringComparison.OrdinalIgnoreCase)
                                    || Convert.ToString(dt.Rows[row - 1][Text]).StartsWith("zzTotal", StringComparison.OrdinalIgnoreCase))
                                {
                                    table.Rows[row].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                                    table.Rows[row].Range.Font.Bold = 1;
                                }
                                table.Cell(row, col).Range.Font.Size = 12.0f;

                                table.Cell(row, col).Height = 22;
                                table.Cell(row, col).Height = 22;
                                //table.Cell(row, col).Width = 5.0f;
                            }
                        }
                        else
                        {
                            GotoastLinePoint();
                            table.Rows[row].Cells.Merge();
                            table.Rows[row].Range.Text = pres_Sortby;
                            table.Rows[row].Range.Font.Size = 14;

                            // table.Rows[row].Height = 25;
                            //table.Rows[row].Cells.Width = 6;
                            table.Rows[row].Range.Font.Color = (WdColor)ColorTranslator.ToOle(Orangecolor);
                            GotoastLinePoint();
                            InserLineBreak();
                            table.Rows[row].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                            GotoastLinePoint();
                            for (int col = 1; col <= columns; col++)
                            {
                                string Text = string.Empty;
                                if (col == 1)
                                {
                                    Text = "ObjectName";
                                }
                                else if (col == 2)
                                {
                                    Text = "PrintCount";
                                }
                                else if (col == 3)
                                {
                                    Text = "OnlineCount";
                                }

                                GotoastLinePoint();


                                table.Cell(row + 1, col).Range.Text = Convert.ToString(dt.Rows[row - 1][Text]);

                                GotoastLinePoint();
                                // Center align text in the cell
                                table.Cell(row + 1, col).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                                // Set vertical alignment for the specified cell
                                table.Cell(row + 1, col).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                                table.Cell(row + 1, col).Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
                                table.Cell(row + 1, col).Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);

                            }
                        }

                    }

                }

                // Add borders to the entire table
                table.Borders.Enable = 1; // Enable borders
                table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle; // Set outside border style
                table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle; // Set inside border style            
                #endregion

                #region Table Style

                // Set border colors
                table.Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
                table.Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);

                #endregion
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void InsertIndexItems(List<NewsArticle> lstNewsArticles)
        {
            Microsoft.Office.Interop.Word.Table objIndextable;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            objIndextable = oDoc.Tables.Add(wrdRng, lstNewsArticles.Count + 1, 4, ref oMissing, ref oMissing);

            objIndextable.Cell(1, 1).Range.Text = "No.";
            objIndextable.Cell(1, 2).Range.Text = "Publication/Portal";
            objIndextable.Cell(1, 3).Range.Text = "Headline";
            objIndextable.Cell(1, 4).Range.Text = "Date";

            objIndextable.Rows[1].Range.Font.Bold = 7;
            objIndextable.Rows[1].Range.Font.ColorIndex = WdColorIndex.wdWhite;
            //objIndextable.Cell(1, 1).Range.Font.Bold = 5;
            //objIndextable.Cell(1, 2).Range.Font.Bold = 5;
            //objIndextable.Cell(1, 3).Range.Font.Bold = 5;
            //objIndextable.Cell(1, 4).Range.Font.Bold = 5;

            objIndextable.Columns[1].PreferredWidth = 35f;
            objIndextable.Columns[2].PreferredWidth = 150f;
            objIndextable.Columns[3].PreferredWidth = 200f;
            objIndextable.Columns[4].PreferredWidth = 80f;



            objIndextable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            Int16 No = 1;

            Microsoft.Office.Interop.Word.Selection wrdSelection = oWord.ActiveWindow.Selection;
            // Microsoft.Office.Interop.Word.Range myRange = mySelection.Range;

            foreach (NewsArticle NewsArticle in lstNewsArticles)
            {
                // Microsoft.Office.Interop.Word.Range myRange = mySelection.Range;
                object linkAddr = "#B" + (No - 1).ToString();
                object Headline = NewsArticle.HeadLine;

                Object oRange = objIndextable.Cell(No + 1, 3).Range;
                wrdSelection.Hyperlinks.Add(oRange, ref linkAddr, ref oMissing, ref Headline, ref Headline, ref oMissing);

                objIndextable.Cell(No + 1, 1).Range.Text = No.ToString();
                objIndextable.Cell(No + 1, 2).Range.Text = NewsArticle.Publication;
                objIndextable.Cell(No + 1, 4).Range.Text = NewsArticle.NewsDate.ToString("MMMM dd, yyyy");


                objIndextable.Cell(No + 1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                objIndextable.Cell(No + 1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                objIndextable.Cell(No + 1, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                objIndextable.Cell(No + 1, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                if ((No % 2) != 0)
                {
                    objIndextable.Rows[No + 1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(lightrowcolor);
                }
                else
                {
                    objIndextable.Rows[No + 1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(darkrowcolor);
                }


                //objIndextable.Rows[No + 1].Range.Borders.InsideColor = WdColor.wdColorWhite;
                objIndextable.Rows[No + 1].Range.Borders.OutsideColor = WdColor.wdColorWhite;

                No++;
            }
            objIndextable.Borders.Enable = 1;
            objIndextable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            objIndextable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;



            objIndextable.Range.Borders.InsideColor = WdColor.wdColorWhite;

            objIndextable.Columns[1].Borders.OutsideColor = WdColor.wdColorWhite;
            objIndextable.Columns[2].Borders.OutsideColor = WdColor.wdColorWhite;
            objIndextable.Columns[3].Borders.OutsideColor = WdColor.wdColorWhite;
            objIndextable.Columns[4].Borders.OutsideColor = WdColor.wdColorWhite;
            objIndextable.Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(Orangecolor);
            objIndextable.Columns[1].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Colomncolor);
            objIndextable.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Orangecolor);


        }

        public void ContentPage()
        {
            try
            {



                Microsoft.Office.Interop.Word.InlineShape oShape;
                object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                //oShape = wrdRng.InlineShapes.AddPicture(URL, ref oMissing, ref oMissing, ref oMissing);



                Microsoft.Office.Interop.Word.Paragraph paragraph = oDoc.Paragraphs.Add();
                //  Microsoft.Office.Interop.Word.Borders borders = paragraph.Borders;

                GotoastLinePoint();
                //paragraph.Range.InsertParagraphAfter();

                //paragraph.Range.Text = "Content";
                //// Explicitly set this to "not bold"
                //paragraph.Range.Font.Bold = 7;
                //paragraph.Range.Font.Size = 15;
                PageHeading_New("Content", "");

                // Set line style to single, color to blue, and line width (adjust as needed)
                //borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                //borders[WdBorderType.wdBorderBottom].Color = (WdColor)ColorTranslator.ToOle(Orangecolor);
                //borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt; // Adjust line width

                // Move to the next line for content
                oWord.Selection.TypeParagraph();

                Microsoft.Office.Interop.Word.Selection wrdSelection = oWord.ActiveWindow.Selection;

                GotoastLinePoint();
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                object linkAddr1 = "#B_OverViewPage";
                object Headline1 = "OverView Page";
                wrdSelection.Hyperlinks.Add(oRng, ref linkAddr1, ref oMissing, ref Headline1, ref Headline1, ref oMissing);

                oWord.ActiveWindow.Selection.InsertAfter("\n");

                GotoastLinePoint();
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                object linkAddr2 = "#B_PageHeadingTable_Print";
                object Headline2 = "Table of Articles-Print";
                wrdSelection.Hyperlinks.Add(oRng, ref linkAddr2, ref oMissing, ref Headline2, ref Headline2, ref oMissing);
                oWord.ActiveWindow.Selection.InsertAfter("\n");

                GotoastLinePoint();
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                object linkAddr3 = "#B_PageHeadingCoverages_Print";
                object Headline3 = "Prominent Coverages-Print";
                wrdSelection.Hyperlinks.Add(oRng, ref linkAddr3, ref oMissing, ref Headline3, ref Headline3, ref oMissing);
                oWord.ActiveWindow.Selection.InsertAfter("\n");

                GotoastLinePoint();
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                object linkAddr4 = "#B_PageHeadingTable_Online";
                object Headline4 = "Table of Articles-Online";
                wrdSelection.Hyperlinks.Add(oRng, ref linkAddr4, ref oMissing, ref Headline4, ref Headline4, ref oMissing);
                oWord.ActiveWindow.Selection.InsertAfter("\n");

                GotoastLinePoint();
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                object linkAddr5 = "#B_PageHeadingCoverages_Online";
                object Headline5 = "Prominent Coverages-Online";
                wrdSelection.Hyperlinks.Add(oRng, ref linkAddr5, ref oMissing, ref Headline5, ref Headline5, ref oMissing);
                oWord.ActiveWindow.Selection.InsertAfter("\n");

            }
            catch (Exception)
            {

                throw;
            }
        }

        public void TextBold(Microsoft.Office.Interop.Word.Table objTable, string text, Int32 Row, Int32 Column)
        {
            Microsoft.Office.Interop.Word.Range cellRange = objTable.Cell(Row, Column).Range;

            // Find the starting index of the content of objNewsArticle.Pageno
            int startIndex = cellRange.Text.IndexOf(text);

            // Set the bold property for the content of objNewsArticle.Pageno
            cellRange.Start = cellRange.Start + startIndex;
            cellRange.End = cellRange.Start + text.Length;
            cellRange.Font.Bold = 1;
        }

        public bool SaveImageFromUrl(string strURL, string strFileName)
        {
            bool IsSuccesful = false;
            string url = "";
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                //string strFetchurl = "https://api.urlbox.io/v1/4G9b0Mqi6yniUmhs/png?url=" + strURL + "&ttl=6000&user_agent=desktop&media=screen&width=600&delay=5000";
                //https://api.urlbox.io/v1/H6TxTXZQLfWwBNl8/png?full_page=true
                //https://api.urlbox.io/v1/DeJyUFW6FSu4yLIt/png?url=https%3A%2F%2Faidossier.adfactorspr.com%2FDossierChart.aspx%3FCDID%3D21%26count%3D1&selector=%23chart1_container&delay=6000

                string strFetchurl = "https://api.urlbox.io/v1/4G9b0Mqi6yniUmhs/png?url=" + HttpUtility.UrlEncode(strURL) + "&selector=%23tblChart&"+Convert.ToString(ConfigurationManager.AppSettings["GraphImageswithHight"])+"";
                strFetchurl = strFetchurl.Replace("\r\n", "");

                //strFetchurl = "https://api.urlbox.io/v1/DeJyUFW6FSu4yLIt/png?url=https%3A%2F%2Faidossier.adfactorspr.com%2FDossierChart.aspx%3FCDID%3D21%26count%3D1&selector=%23chart1_container&delay=6000";
                //4G9b0Mqi6yniUmhs
                WebRequest request = WebRequest.Create(strFetchurl);

                // Get the response.
                WebResponse response = request.GetResponse();

                // Get the stream containing content returned by the server.
                Stream dataStream = response.GetResponseStream();

                // Read the content.
                System.Drawing.Image image = System.Drawing.Image.FromStream(dataStream);

                //@"D:\\ScreenshotImage\\img_" + counter + "_" + DateTime.Now.ToFileTime()

                string strPath = System.Configuration.ConfigurationManager.AppSettings["ScreenshotSavePath"] + strFileName + ".jpg";//ConfigurationManager.AppSettings["ScreenshotExtension"];
                image.Save(strPath);



                if (strPath != null)
                {
                    //string BookMark1 = "";

                    InsertDomainImageIntoBody_New(strPath);
                    //ErrorLog("GenerateDossier", i + " Article Image added");

                }

                // lstInput.Where(w => w.URL == strURL).ToList().ForEach(i => i.ScreenshotUrl = strPath);
                //Delete file 
                if (strPath != null || strPath != string.Empty)
                {
                    if ((System.IO.File.Exists(strPath)))
                    {
                        System.IO.File.Delete(strPath);
                    }

                }
                IsSuccesful = true;
            }
            catch (Exception ex)
            {
                //ErrorLog("SaveImageFromUrl", ex.Message + " ->> " + ex.StackTrace + " ->> " + url + " ->> " + strURL);
            }

            return IsSuccesful;
        }
        public bool SaveImageFromUrl_4_Graph(string strURL, string strFileName)
        {
            bool IsSuccesful = false;
            string url = "";
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                //string strFetchurl = "https://api.urlbox.io/v1/4G9b0Mqi6yniUmhs/png?url=" + strURL + "&ttl=6000&user_agent=desktop&media=screen&width=600&delay=5000";
                //https://api.urlbox.io/v1/H6TxTXZQLfWwBNl8/png?full_page=true
                //https://api.urlbox.io/v1/DeJyUFW6FSu4yLIt/png?url=https%3A%2F%2Faidossier.adfactorspr.com%2FDossierChart.aspx%3FCDID%3D21%26count%3D1&selector=%23chart1_container&delay=6000

                string strFetchurl = "https://api.urlbox.io/v1/4G9b0Mqi6yniUmhs/png?url=" + HttpUtility.UrlEncode(strURL) + "&selector=%23tblChart&"+Convert.ToString(ConfigurationManager.AppSettings["GraphImageswithHight_4"])+"";
                strFetchurl = strFetchurl.Replace("\r\n", "");

                //strFetchurl = "https://api.urlbox.io/v1/DeJyUFW6FSu4yLIt/png?url=https%3A%2F%2Faidossier.adfactorspr.com%2FDossierChart.aspx%3FCDID%3D21%26count%3D1&selector=%23chart1_container&delay=6000";
                //4G9b0Mqi6yniUmhs
                WebRequest request = WebRequest.Create(strFetchurl);

                // Get the response.
                WebResponse response = request.GetResponse();

                // Get the stream containing content returned by the server.
                Stream dataStream = response.GetResponseStream();

                // Read the content.
                System.Drawing.Image image = System.Drawing.Image.FromStream(dataStream);

                //@"D:\\ScreenshotImage\\img_" + counter + "_" + DateTime.Now.ToFileTime()

                string strPath = System.Configuration.ConfigurationManager.AppSettings["ScreenshotSavePath"] + strFileName + ".jpg";//ConfigurationManager.AppSettings["ScreenshotExtension"];
                image.Save(strPath);



                if (strPath != null)
                {
                    //string BookMark1 = "";

                    InsertDomainImageIntoBody_New_4Graph(strPath);
                    //ErrorLog("GenerateDossier", i + " Article Image added");

                }

                // lstInput.Where(w => w.URL == strURL).ToList().ForEach(i => i.ScreenshotUrl = strPath);
                //Delete file 
                if (strPath != null || strPath != string.Empty)
                {
                    if ((System.IO.File.Exists(strPath)))
                    {
                        System.IO.File.Delete(strPath);
                    }

                }
                IsSuccesful = true;
            }
            catch (Exception ex)
            {
                //ErrorLog("SaveImageFromUrl", ex.Message + " ->> " + ex.StackTrace + " ->> " + url + " ->> " + strURL);
            }

            return IsSuccesful;
        }
        


        public bool SaveImageFromUrl_Online(string strURL, string strFileName, string BookMark)
        {
            bool IsSuccesful = false;

            try
            {

                Microsoft.Office.Interop.Word.InlineShape oShape;
                object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;


                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                //string strFetchurl = "https://api.urlbox.io/v1/4G9b0Mqi6yniUmhs/png?url=" + strURL + "&ttl=6000&user_agent=desktop&media=screen&width=600&delay=5000";
                //https://api.urlbox.io/v1/H6TxTXZQLfWwBNl8/png?full_page=true
                //https://api.urlbox.io/v1/DeJyUFW6FSu4yLIt/png?url=https%3A%2F%2Faidossier.adfactorspr.com%2FDossierChart.aspx%3FCDID%3D21%26count%3D1&selector=%23chart1_container&delay=6000

                string strFetchurl = "https://api.urlbox.io/v1/4G9b0Mqi6yniUmhs/png?url=" + HttpUtility.UrlEncode(strURL) + "&selector=%23tblChart&delay=6000";
                strFetchurl = strFetchurl.Replace("\r\n", "");

                //strFetchurl = "https://api.urlbox.io/v1/DeJyUFW6FSu4yLIt/png?url=https%3A%2F%2Faidossier.adfactorspr.com%2FDossierChart.aspx%3FCDID%3D21%26count%3D1&selector=%23chart1_container&delay=6000";
                //4G9b0Mqi6yniUmhs
                WebRequest request = WebRequest.Create(strFetchurl);




                //    // Get the response. 
                WebResponse response = request.GetResponse();


                // Get the stream containing content returned by the server.
                Stream dataStream = response.GetResponseStream();
                System.Drawing.Image image = System.Drawing.Image.FromStream(dataStream);
                string strPath = System.Configuration.ConfigurationManager.AppSettings["ScreenshotSavePath_Online"] + strFileName + ".jpg";//ConfigurationManager.AppSettings["ScreenshotExtension"];
                image.Save(strPath);

                if (strPath != null)
                {
                    //string BookMark1 = "";

                    InsertDomainImageIntoBody_New(strPath);
                    //ErrorLog("GenerateDossier", i + " Article Image added");

                }

                
                //BOOKMARKING
                if (!string.IsNullOrEmpty(BookMark))
                {
                    oDoc.Bookmarks.Add(BookMark, oMissing);
                }
                //Delete file 
                if (strPath != null || strPath != string.Empty)
                {
                    if ((System.IO.File.Exists(strPath)))
                    {
                        System.IO.File.Delete(strPath);
                    }
                }

                // lstInput.Where(w => w.URL == strURL).ToList().ForEach(i =;> i.ScreenshotUrl = strPath);

                IsSuccesful = true;
            }

            catch (Exception)
            {
                //       word.InlineShape inlineShape = oWord.Selection.InlineShapes.AddPicture(ConfigurationManager.AppSettings["Imagenotavailable"], oMissing, oMissing, oMissing);

                word.InlineShape inlineShape = oWord.Selection.InlineShapes.AddPicture(ConfigurationManager.AppSettings["Imagenotavailable"], oMissing, oMissing, oMissing);
                inlineShape.Height = 270.0f;
                // throw ;
                //ErrorLog("SaveImageFromUrl", ex.Message + " ->> " + ex.StackTrace + " ->> " + url + " ->> " + strURL);
            }



            return IsSuccesful;
        }
    }
}
