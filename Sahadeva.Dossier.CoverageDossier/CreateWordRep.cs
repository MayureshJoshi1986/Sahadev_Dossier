using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Data;

namespace Sahadeva.Dossier.CoverageDossier
{
     class CreateWordRep
    {
        public  void Main1()
        {
            // Create a new instance of the Word application
            Application wordApp = new Application();
            
            // Path to the document
            object filePath = @"D:\ADFACTORS\TFS Applications\PR\Workbench\Dossier Template2\WordPOC\test.docx";
            // Create a new instance of the Word application
           

           
            object missing = Type.Missing;

            // Open the document
            Document doc = wordApp.Documents.Open(ref filePath, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            // Make the application visible (optional)
            wordApp.Visible = true;

            // Find and replace text
            FindAndReplace(wordApp, "#Clientname", "TCS (Tata Consultancy Services)");
            FindAndReplace(wordApp, "#TopicDetails", "TCS Q4 results");
            FindAndReplace(wordApp, "#Date", "10th April 2024 to 19th April 2024");

            GotoastLinePoint(wordApp);
            InsertPagebreak(wordApp);
            // Add a new page
           // AddNewPage(wordApp);
            Dossier.DAL.DossierTemp1DAL dalDT1 = new DAL.DossierTemp1DAL();
            DataSet dsOverViewPage = dalDT1.CoverageDossier_OverView_Page_DT1(7781);

            // Add a new page and insert a table
            //AddNewPageAndTable(wordApp);
            //// Save and close the document
            //doc.Save();
            //doc.Close();

            // Path to save the new document
            object newFilePath = @"D:\ADFACTORS\TFS Applications\PR\Workbench\Dossier Template2\WordPOC\" + "TCS (Tata Consultancy Services)" + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + ".docx";

            // Save the document with a new name
            doc.SaveAs(ref newFilePath, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            // Close the document
            //doc.Close(ref missing, ref missing, ref missing);

            // Quit the Word application
            wordApp.Quit();

            Console.WriteLine("Text replaced and document saved as new file successfully!");
       
        }

        static void FindAndReplace(Application wordApp, object findText, object replaceText)
        {
            // Set up the find and replace parameters
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Text = (string)findText;
            wordApp.Selection.Find.Replacement.ClearFormatting();
            wordApp.Selection.Find.Replacement.Text = (string)replaceText;

            // Execute the find and replace operation
            object replaceAll = WdReplace.wdReplaceAll;
            wordApp.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceText, ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }

       // private static object missing = System.Reflection.Missing.Value;
    
         static void AddNewPage(Application wordApp)
        {
            // Move the selection to the end of the document
            object collapseEnd = WdCollapseDirection.wdCollapseEnd;
            wordApp.Selection.Collapse(ref collapseEnd);

            // Insert a page break
            object breakType = WdBreakType.wdPageBreak;
            wordApp.Selection.InsertBreak(ref breakType);
        }
         public void GotoastLinePoint(Application wordApp)
         {
             object what = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPercent;
             object which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToLast;
             wordApp.Selection.GoTo(ref what, ref which, missing, missing);
         }

         public void InsertPagebreak(Application wordApp)
         {
             // VB : Selection.InsertBreak Type:=wdPageBreak
             object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
             wordApp.Selection.InsertBreak(ref pBreak);
             //oWord.ActiveWindow.Selection.InsertBreak(ref pBreak);
             //oWord.ActiveWindow.Selection.InsertAfter("\n");
         }
         public void InserLineBreak(Application wordApp)
         {
             object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak;
             wordApp.Selection.InsertBreak(ref pBreak);
         }
         static void AddNewPageAndTable(Application wordApp)
        {
            // Move the selection to the end of the document
            object collapseEnd = WdCollapseDirection.wdCollapseEnd;
            wordApp.Selection.Collapse(ref collapseEnd);

            // Insert a page break
            object breakType = WdBreakType.wdPageBreak;
            wordApp.Selection.InsertBreak(ref breakType);

            // Insert a table on the new page
            Range range = wordApp.Selection.Range;
            Table table = wordApp.ActiveDocument.Tables.Add(range, 3, 3, ref missing, ref missing);

            // Fill the table with sample data
            for (int r = 1; r <= 3; r++)
            {
                for (int c = 1; c <= 3; c++)
                {
                    table.Cell(r, c).Range.Text = "Row {r}, Col {c}";
                }
            }
        }
         public static void OverViewPage(DataSet dsOverViewPage, Application wordApp)
         {
             try
             {
                 if (dsOverViewPage != null)
                 {
                     if (dsOverViewPage.Tables.Count > 0)
                     {
                         for (int i = 0; i < dsOverViewPage.Tables.Count; i++)
                         {
                             if (dsOverViewPage.Tables[i] != null)
                             {
                                 if (dsOverViewPage.Tables[i].Rows.Count > 0)
                                 {

                                     //wordApp.PageHeading_New("Overview", "");

                                     //WordAutomationObjTemp1.InserLineBreak();

                                     //string objBookmark = "B_OverViewPage";


                                     //ErrorLog("Inside CreateSimpleTable_1", "Started");

                                     //WordAutomationObjTemp1.CreateSimpleTable_1(dsOverViewPage.Tables[i].Rows.Count, 3, dsOverViewPage.Tables[i], objBookmark);

                                     //ErrorLog("Inside CreateSimpleTable_1", "End");

                                 }

                             }

                         }

                     }

                 }

             }
             catch (Exception)
             {

                 throw;
             }
         }

         public void PageHeading_New(string Text, string objBookmark)
         {


             //Microsoft.Office.Interop.Word.Table table1;
             //Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
             //table1 = oDoc.Tables.Add(wrdRng, 1, 1, ref oMissing, ref oMissing);
             //table1.Range.ParagraphFormat.SpaceAfter = 1;

             ////BOOKMARKING
             //if (!string.IsNullOrEmpty(objBookmark))
             //{
             //    oDoc.Bookmarks.Add(objBookmark, wrdRng);
             //}

             //table1.Cell(1, 1).Range.Text = Text;
             //table1.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

             //// Set vertical alignment for the specified cell
             //table1.Cell(1, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;


             //table1.Cell(1, 1).Range.Borders.InsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
             //table1.Cell(1, 1).Range.Borders.OutsideColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);

             //table1.Rows[1].Range.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(FadedOrangecolor);
             //table1.Rows[1].Range.Font.Bold = 600;
             //table1.Cell(1, 1).Range.Font.Size = 15.0f;

             //table1.Cell(1, 1).Height = 22;
             //table1.Cell(1, 1).Height = 22;
             //GotoastLinePoint();




         }
        private static object missing = System.Reflection.Missing.Value;
    }
}
    

