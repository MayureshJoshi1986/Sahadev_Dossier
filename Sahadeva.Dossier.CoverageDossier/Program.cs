using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sahadeva.Dossier.Common;
using System.IO;
using System.Configuration;
using Sahadeva.Dossier.Entities;
using System.Data;
using System.Net.Mail;
using System.Web;
using System.Net;
using System.Diagnostics;
using System.Threading.Tasks;



namespace Sahadeva.Dossier.CoverageDossier
{
    public class Program
    {


        static void Main(string[] args)
        {
            string ClientEmailIdsTo = "";
            string EntityName = "";
            string EmailIdsCC = " ";
            string EmailIdsBCC = "";

            try
            {





                Dossier.DAL.dossierDAL dal = new DAL.dossierDAL();

                DataTable lstCDID = dal.FetchPending_DCIDsToProcess_All();

                //// //lstCDID.Rows.Add(7689);
                //// DT2.MailDT2(lstCDID);
                ////// DataTable lstCDID = new DataTable();

                DossierTemp1 DT1 = new DossierTemp1();
                DossierTemp2 DT2 = new DossierTemp2();
                DossierTemp3 DT3 = new DossierTemp3();
                DossierTemp4 DT4 = new DossierTemp4();

                DataTable lstCDIDDT1 = new DataTable();
                DataTable lstCDIDDT2 = new DataTable();
               // DataTable lstCDIDDT3 = new DataTable();
                DataTable lstCDIDDT4 = new DataTable();


               // lstCDID.Rows.Add(7808, 1, 4);

                //lstCDID.Rows.Add(7818, 1, 1);
                lstCDID.Rows.Add(7809, 1, 2);

                //  lstCDID.Rows.Add(12, 1, 3);
                //// DT3.MailDT3(lstCDID);

                if (lstCDID.Rows.Count > 0)
                {

                    for (int i = 0; i < lstCDID.Rows.Count; i++)
                    {
                        if (lstCDID.Rows[i]["TemplateID"].ToString() == "1")
                        {                         

                            // Check if lstCDIDDT1 has a column named "CDID"
                            if (!lstCDIDDT1.Columns.Contains("CDID"))
                            {
                                lstCDIDDT1.Columns.Add("CDID", typeof(int));
                            }

                            // Add the value to the new DataTable
                            lstCDIDDT1.Rows.Add(Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                        }
                        else if (lstCDID.Rows[i]["TemplateID"].ToString() == "2")
                        {                            
                            // Check if lstCDIDDT1 has a column named "CDID"
                            if (!lstCDIDDT2.Columns.Contains("CDID"))
                            {
                                lstCDIDDT2.Columns.Add("CDID", typeof(int));
                            }

                            // Add the value to the new DataTable
                            lstCDIDDT2.Rows.Add(Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                        }
                  
                        else if (lstCDID.Rows[i]["TemplateID"].ToString() == "4")
                        {
                            if (!lstCDIDDT4.Columns.Contains("CDID"))
                            {
                                lstCDIDDT4.Columns.Add("CDID", typeof(int));
                            }

                            // Add the value to the new DataTable
                            lstCDIDDT4.Rows.Add(Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                        }

                    }


                    var stopWatch = Stopwatch.StartNew();
                    ParallelOptions po = new ParallelOptions();
                    po.MaxDegreeOfParallelism = 1;
                    System.Net.ServicePointManager.DefaultConnectionLimit = 1; Parallel.For(0, 1, po, i =>
                    {
                        try
                        {
                            DT1.MainDT1(lstCDIDDT1);
                             DT2.MainDT2(lstCDIDDT2);
                          
                        }
                        catch (Exception ex)
                        {
                        }
                    });
                    double time = stopWatch.Elapsed.TotalSeconds / 60;



                  
                }
            }
            catch (Exception ex)
            {
                ErrorLog("Error Log", ex.ToString());
                if (ClientEmailIdsTo != "")
                {
                    //ErrorSendMail(ClientEmailIdsTo, "No Articles Data Found", "Dossier Generator - " + EntityName, "", EmailIdsCC, EmailIdsBCC);
                }
                //  ErrorSendMail(Convert.ToString(ConfigurationManager.AppSettings["ErrorEmailID"]), ex.ToString(), "Error Log Dossier Generator - " + EntityName, "", EmailIdsCC, EmailIdsBCC);
                // throw;
                //}




                //Test(ds);
                //Test_1(ds);



            }
        }

        
        public static void CreateWord(List<NewsArticle> lstNewsArticle, List<NewsArticle_Online> lstNewsArticleOnline, DataTable dtClient, Int32 CDID)
        {
            try
            {

                WordAutomation WordAutomationObj = new WordAutomation();
               
                string BookMark = string.Empty;
                //WordAutomationObj.InsertDocumentHeadline("Table of Contents-Print");

                Dossier.DAL.dossierDAL dal = new DAL.dossierDAL();


                WordAutomationObj.GotoastLinePoint();
                string FirstpagebgURL = ConfigurationManager.AppSettings["Firstpagebg"];
                //WordAutomationObj.InsertDomainImageIntoBody_New(FirstpagebgURL);
               // WordAutomationObj.InsertImageAndWriteTextOnIt(FirstpagebgURL, Convert.ToString(dtClient.Rows[0]["EntityName"]));



                WordAutomationObj.FirstPage_New(Convert.ToString(dtClient.Rows[0]["EntityName"]), Convert.ToString(dtClient.Rows[0]["Description"]), Convert.ToString(dtClient.Rows[0]["FromDate"]) + " to " + Convert.ToString(dtClient.Rows[0]["ToDate"]), FirstpagebgURL);
                WordAutomationObj.InsertPagebreak();

                //WordAutomationObj.GotoastLinePoint();
                //WordAutomationObj.ContentPage();
                //WordAutomationObj.InsertPagebreak();

               // WordAutomationObj.GotoastLinePoint();
                //WordAutomationObj.InsertPagebreak();

                //WordAutomationObj.InsertIndexItems(lstNewsArticle);

                WordAutomationObj.GotoastLinePoint();
                 WordAutomationObj.InserLineBreak();

                 DataSet dsOverViewPage = dal.CoverageDossier_OverView_Page_DT2(CDID);

                 OverViewPage(dsOverViewPage, WordAutomationObj);            //Overview Page

                 
                //WordAutomationObj.GotoastLinePoint();
                //WordAutomationObj.InserLineBreak();
                WordAutomationObj.PageHeading_New("Overview summary", "");
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();

                WordAutomationObj.InsertText(Convert.ToString(dtClient.Rows[0]["Summary"]));

                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
                WordAutomationObj.InsertPagebreak();

                PrintGraph(CDID, WordAutomationObj);        // Overview Print Graph 

                WordAutomationObj.InsertPagebreak();

                
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
                ErrorLog("Article vs Reach Graph", "Graph Article vs Reach Start");
                ArticleReach_Graph(CDID, WordAutomationObj);        //  Graph  Coverage Category
                ErrorLog("Article vs Reach Graph", "Graph  Article vs Reach End");
                WordAutomationObj.InsertPagebreak();



               
                 //WordAutomationObj.InsertPagebreak();

                ErrorLog("Inside CreateWord - Print", "Started");

                #region For Print

                //WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();

               

                if (lstNewsArticle.Count != 0)
                {
                    //WordAutomationObj.PageHeading("Table of Contents-Print", "B_PageHeadingTable_Print");
                    WordAutomationObj.PageHeading_New("Table of contents - print", "B_PageHeadingTable_Print");
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();


                    WordAutomationObj.InsertTable_Articles_Print(lstNewsArticle);
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();
                    WordAutomationObj.InsertPagebreak();


                    // WordAutomationObj.PageHeading("Prominent Coverages", "B_PageHeadingCoverages_Print");
                    WordAutomationObj.PageHeading_New("Prominent Coverages – Print", "B_PageHeadingCoverages_Print");
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();

                    for (int i = 0; i < lstNewsArticle.Count; i++)
                    {
                        ErrorLog("Inside CreateWord - Print ", "Started lstNewsArticle Count : " + Convert.ToString(i));
                        NewsArticle objNewsArticle = new NewsArticle();

                        objNewsArticle = lstNewsArticle[i];

                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();

                        //WordAutomationObj.InsertDomainImageIntoBody_New(@"D:\Shivakiran\Backup_Prashant\Sahadeva_NewsLetter\Sahadeva.Dossier.CoverageDossier\Images\Publications\110.jpg");

                        WordAutomationObj.InsertHeaderTable_New_Print(objNewsArticle);

                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();

                        //WordAutomationObj.InsertDocumentHeadline(Convert.ToString(objNewsArticle.HeadLine));
                        //WordAutomationObj.GotoastLinePoint();



                        //WordAutomationObj.InsertDomainImageIntoBody("https://upload.wikimedia.org/wikipedia/en/thumb/9/9a/ONGC_Logo.svg/200px-ONGC_Logo.svg.png");
                        //ErrorLog("GenerateDossier", i + " Publication Logo added");
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        //}
                        WordAutomationObj.InsertBoldTextLable("Article synopsis - ");
                        WordAutomationObj.InsertText(Convert.ToString(objNewsArticle.NewsText));
                        WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InsertText_WithoutNewLine(Convert.ToString(objNewsArticle.NewsDate.ToString("dd MMMM, yyyy")) + " | " + Convert.ToString(objNewsArticle.Edition) + " Edition | ");
                        //WordAutomationObj.InsertHyperLink("Article Link", Convert.ToString(objNewsArticle.NewsURL));
                        //WordAutomationObj.GotoastLinePoint();

                        //WordAutomationObj.InsertHeaderTable_New_Print(objNewsArticle);
                        //ErrorLog("GenerateDossier", i + " Header table added");
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();




                        //WordAutomationObj.InsertNewsArticleImageIntoBody("NewsImage", @"D:\images\What-is-an-article.jpg");
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();

                        //string BookMark = "B" + 1;
                        string ArticleIamgeURL = ConfigurationManager.AppSettings["ArticleIamgeURL"];

                        if (objNewsArticle.NewsURL != null)
                        {
                            BookMark = "B_Print" + lstNewsArticle.IndexOf(lstNewsArticle[i]).ToString();

                            WordAutomationObj.InsertArticleImageIntoBody(ArticleIamgeURL + objNewsArticle.IMG_URL, BookMark);
                            //ErrorLog("GenerateDossier", i + " Article Image added");
                            WordAutomationObj.GotoastLinePoint();
                        }
                        //WordAutomationObj.InsertArticleImageIntoBody(ArticleIamgeURL + objNewsArticle.IMG_URL, BookMark);
                        //ErrorLog("GenerateDossier", i + " Article Image added");
                        // WordAutomationObj.GotoastLinePoint();

                        //WordAutomationObj.GotoastLinePoint();
                        // WordAutomationObj.InserLineBreak();
                        WordAutomationObj.InsertPagebreak();

                        ////KillWordProcess();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InsertPagebreak();
                        ErrorLog("Inside CreateWord - Print ", "End lstNewsArticle Count : " + Convert.ToString(i));

                    }

                }
                //WordAutomationObj.GotoastLinePoint();
                //WordAutomationObj.InserLineBreak();
                string EntityID = Convert.ToString(dtClient.Rows[0]["EntityID"]);
                string ClientLogoURL = Convert.ToString(ConfigurationManager.AppSettings["ClientLogoURL"]);
                string ADFLogoURL = Convert.ToString(ConfigurationManager.AppSettings["ADFLogoURL"]);

                //Logo check 
                if (EntityID != null || EntityID != string.Empty)
                {
                    if ((System.IO.File.Exists(ClientLogoURL + EntityID + ".jpg")))
                    {
                        WordAutomationObj.InsertHeaderandFooter(ClientLogoURL + EntityID + ".jpg", ADFLogoURL);
                    }
                    else
                    {
                        WordAutomationObj.InsertHeaderandFooter(ADFLogoURL, ADFLogoURL);
                    }
                }

                //  WordAutomationObj.InsertHeaderandFooter(ClientLogoURL + EntityID + ".jpg", ADFLogoURL);
                WordAutomationObj.GotoastLinePoint();
                ////WordAutomationObj.InsertPagebreak();
                #endregion

                ErrorLog("Inside CreateWord - Print", "End");
                ErrorLog("Inside CreateWord - Online", "Started");

                #region For Online

                //WordAutomationObj.GotoastLinePoint();
                //WordAutomationObj.InserLineBreak();

                

                if (lstNewsArticleOnline.Count != 0)
                {
                    //  WordAutomationObj.PageHeading("Table of Contents-Online", "B_PageHeadingTable_Online");
                    WordAutomationObj.PageHeading_New("Table of contents - online", "B_PageHeadingTable_Online");
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();

                    WordAutomationObj.InsertTable_Articles_Online(lstNewsArticleOnline);
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();
                    WordAutomationObj.InsertPagebreak();



                    WordAutomationObj.PageHeading_New("Prominent Coverages – Online", "B_PageHeadingCoverages_Online");
                    // WordAutomationObj.PageHeading("Prominent Coverages", "B_PageHeadingCoverages_Online");
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();

                    for (int i = 0; i < lstNewsArticleOnline.Count; i++)
                    {
                        ErrorLog("Inside CreateWord - Online ", "Started lstNewsArticle Count : " + Convert.ToString(i));
                        NewsArticle_Online objNewsArticle = new NewsArticle_Online();

                        objNewsArticle = lstNewsArticleOnline[i];

                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();

                        //WordAutomationObj.InsertDomainImageIntoBody_New(@"D:\Shivakiran\Backup_Prashant\Sahadeva_NewsLetter\Sahadeva.Dossier.CoverageDossier\Images\Publications\110.jpg");

                        WordAutomationObj.InsertHeaderTable_New_Online(objNewsArticle);

                        WordAutomationObj.GotoastLinePoint();
                        //  WordAutomationObj.InserLineBreak();

                        WordAutomationObj.InsertDocumentHeadline(Convert.ToString(objNewsArticle.Title));
                        WordAutomationObj.GotoastLinePoint();



                        //WordAutomationObj.InsertDomainImageIntoBody("https://upload.wikimedia.org/wikipedia/en/thumb/9/9a/ONGC_Logo.svg/200px-ONGC_Logo.svg.png");
                        //ErrorLog("GenerateDossier", i + " Publication Logo added");
                        WordAutomationObj.GotoastLinePoint();
                        // WordAutomationObj.InserLineBreak();
                        //}
                        WordAutomationObj.InsertBoldTextLable("Article synopsis - ");
                        WordAutomationObj.InsertText(Convert.ToString(objNewsArticle.NewsText));
                        WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InsertText_WithoutNewLine(Convert.ToString(objNewsArticle.NewsDate.ToString("dd MMMM, yyyy")) + " | " + Convert.ToString(objNewsArticle.Edition) + " Edition | ");
                        //WordAutomationObj.InsertHyperLink("Article Link", Convert.ToString(objNewsArticle.NewsURL));
                        //WordAutomationObj.GotoastLinePoint();

                        //WordAutomationObj.InsertHeaderTable_New_Online(objNewsArticle);
                        //ErrorLog("GenerateDossier", i + " Header table added");
                        WordAutomationObj.GotoastLinePoint();
                        //  WordAutomationObj.InserLineBreak();




                        //WordAutomationObj.InsertNewsArticleImageIntoBody("NewsImage", @"D:\images\What-is-an-article.jpg");
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();


                        string ArticleIamgeURL = ConfigurationManager.AppSettings["ArticleOnlineIamgeURL"];


                        //BookMark = "B_Online" + lstNewsArticleOnline.IndexOf(lstNewsArticleOnline[i]).ToString();

                        //if (objNewsArticle.NewsURL != null)
                        //{
                        //    WordAutomationObj.SaveImageFromUrl_Online(objNewsArticle.NewsURL, (CDID + "_Online_" + i).ToString(), BookMark);
                        //    WordAutomationObj.GotoastLinePoint();
                        //}
                        if (objNewsArticle.screenshot_fileName != null)
                        {
                            BookMark = "B_Online" + lstNewsArticleOnline.IndexOf(lstNewsArticleOnline[i]).ToString();
                            string ext = Path.GetExtension(ArticleIamgeURL + objNewsArticle.screenshot_fileName);
                            if (ext != null)
                            {
                                //ErrorLog("objNewsArticle.IMG_URL", objNewsArticle.IMG_URL.ToString());
                                //  WordAutomationObj.InsertArticleImageIntoBody(objNewsArticle.IMG_URL, BookMark);
                                WordAutomationObj.InsertArticleImageIntoBody(ArticleIamgeURL + objNewsArticle.screenshot_fileName, BookMark);
                            }
                            else
                            {
                                WordAutomationObj.InsertArticleImageIntoBody(ArticleIamgeURL + objNewsArticle.screenshot_fileName + ".jpg", BookMark);
                            }
                            //  WordAutomationObj.SaveImageFromUrl_Online(objNewsArticle.NewsURL, (CDID + "_Online_" + i).ToString(), BookMark);
                            //ErrorLog("GenerateDossier", i + " Article Image added");
                            //  WordAutomationObj.GotoastLinePoint();
                        }
                        //WordAutomationObj.InsertArticleImageIntoBody(objNewsArticle.IMG_URL, BookMark);
                        //ErrorLog("GenerateDossier", i + " Article Image added");
                        //WordAutomationObj.GotoastLinePoint();

                        //WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InsertPagebreak();

                        ////KillWordProcess();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InsertPagebreak();
                        ErrorLog("Inside CreateWord - Print ", "Started lstNewsArticle Count : " + Convert.ToString(i));
                    }
                }

                //WordAutomationObj.InsertDomainImageIntoBody("data:image/jpeg;base64,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");
                //WordAutomationObj.GotoastLinePoint();
                //WordAutomationObj.InserLineBreak();

                //WordAutomationObj.InsertHeaderandFooter_New(@"D:\Shivakiran\Backup_Prashant\Sahadeva_NewsLetter\Sahadeva.Dossier.CoverageDossier\Images\Publications\9.jpg", @"D:\Shivakiran\Backup_Prashant\Sahadeva_NewsLetter\Sahadeva.Dossier.CoverageDossier\Images\adflogo.jpg");
                //WordAutomationObj.GotoastLinePoint();
                //WordAutomationObj.InsertPagebreak();
                #endregion

                ErrorLog("Inside CreateWord - Online", "End");
               // WordAutomationObj.GotoastLinePoint();
               // WordAutomationObj.InserLineBreak();
                //WordAutomationObj.InsertPagebreak();

                WordAutomationObj.PageHeading_New("Brand vs Competitor Overview-Coverage Category", "B_PageHeading_BrandCompetitor_Overview");
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
                ErrorLog("Brand vs Competitor overview", "Graph  Coverage Category Start");
                CoverageCategory_Graph(CDID, WordAutomationObj);        //  Graph  Coverage Category
                ErrorLog("Brand vs Competitor overview", "Graph  Coverage Category End");

                WordAutomationObj.PageHeading_New("Brand vs Competitor Overview-Article Type", "B_PageHeading_BrandCompetitor_Overview");
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
                ErrorLog("Brand vs Competitor overview", "Graph  Article TYPE Graph  Start");
                Article_TYPE_Graph(CDID, WordAutomationObj);        // Article TYPE Graph 
                ErrorLog("Brand vs Competitor overview", "Graph  Article TYPE Graph  End");
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
                WordAutomationObj.InsertPagebreak();

                WordAutomationObj.PageHeading_New("Mention Comparison", "B_PageHeading_BrandCompetitor_Overview");

                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
                ErrorLog("Brand vs Competitor overview", "Graph  Mention_Graph Graph   Start");
                Mention_TYPE_Graph(CDID, WordAutomationObj);        // AMention_Graph Graph 
                ErrorLog("Brand vs Competitor overview", "Graph  Mention_Graph Graph   end");
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();


                WordAutomationObj.PageHeading_New("Brand vs Competitor Publication Type", "B_PageHeading_BrandCompetitor_Overview");
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
               ErrorLog("Brand vs Competitor overview", "Graph  PublicationType Graph Start");
                
                //PublicationType_Graph(CDID, WordAutomationObj);
              
                PublicationType_Graph1(CDID, WordAutomationObj);   
                ErrorLog("Brand vs Competitor overview", "Graph  PublicationType Graph end");
              
                

              


                String DossierFileName = Convert.ToString(dtClient.Rows[0]["EntityName"]).Replace(" ", "_").Trim() + "_" +
                                 DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString()
                                + "_" + Convert.ToString(dtClient.Rows[0]["FromDate"]).Replace(" ", "_") +
                                "_" + Convert.ToString(dtClient.Rows[0]["ToDate"]).Replace(" ", "_")
                                + Convert.ToString(dtClient.Rows[0]["CDID"]);

                //string SavePath = Convert.ToString(dtClient.Rows[0]["EntityID"]);
                string SavePath = Convert.ToString(ConfigurationManager.AppSettings["SavePath"]);
                WordAutomationObj.SaveAs(SavePath + DossierFileName + ".doc");
                WordAutomationObj.oWord.Quit();
                string OutputFiles_URL = Convert.ToString(ConfigurationManager.AppSettings["OutputFiles_URL"]);

                string EmailBody = System.IO.File.ReadAllText(ConfigurationManager.AppSettings["MailBodypath"].ToString());

                EmailBody = EmailBody.Replace("#DossierNo#", "PD-" + Convert.ToString(dtClient.Rows[0]["CDCID"]));
                EmailBody = EmailBody.Replace("#link#", OutputFiles_URL + DossierFileName + ".doc");
                EmailBody = EmailBody.Replace("#EntityName#", Convert.ToString(dtClient.Rows[0]["EntityName"]));
                EmailBody = EmailBody.Replace("#EmpName#", Convert.ToString(dtClient.Rows[0]["Name"])); 
                
                   // "Your request for Dossier" + Convert.ToString(dtClient.Rows[0]["EntityName"]) + " is ready to be downloaded. For downloading the dossier, please click on the link below are go to the download section in the Dossier platform. " + OutputFiles_URL + DossierFileName + ".doc";

              dal.Update_CD_OutputFileLink(Convert.ToInt32(dtClient.Rows[0]["CDID"]), OutputFiles_URL + DossierFileName + ".doc");

             // SendMail(Convert.ToString(dtClient.Rows[0]["EmailIdsTo"]), EmailBody, "Dossier Generator - " + Convert.ToString(dtClient.Rows[0]["EntityName"]), "", Convert.ToString(dtClient.Rows[0]["EmailIdsCC"]), Convert.ToString(dtClient.Rows[0]["EmailIdsBCC"] + "," + ConfigurationManager.AppSettings["ErrorEmailID"]));
            }
            catch (Exception)
            {

                  //ErrorSendMail(Convert.ToString(dtClient.Rows[0]["EmailIdsTo"]), EmailBody, "Dossier Generator - " + Convert.ToString(dtClient.Rows[0]["EntityName"]), "", Convert.ToString(dtClient.Rows[0]["EmailIdsCC"]), Convert.ToString(dtClient.Rows[0]["EmailIdsBCC"]));

                throw;
            }
        }


        public static void Test(DataSet ds)
        {
            WordAutomation WordAutomationObj = new WordAutomation();

            if (ds != null)
            {
                if (ds.Tables.Count > 0)
                {
                    for (int i = 0; i < ds.Tables.Count; i++)
                    {
                        if (ds.Tables[i] != null)
                        {
                            if (ds.Tables[i].Rows.Count > 0)
                            {
                                WordAutomationObj.GotoastLinePoint();
                                WordAutomationObj.AddColspanEffect(1, 3, Convert.ToString(ds.Tables[i].Rows[0]["Sortby"]));
                                WordAutomationObj.GotoastLinePoint();
                                WordAutomationObj.NextLine();
                                WordAutomationObj.GotoastLinePoint();
                                //WordAutomationObj.GotoastLinePoint();
                                //WordAutomationObj.InserLineBreak();

                                WordAutomationObj.CreateSimpleTable(ds.Tables[i].Rows.Count, 3, ds.Tables[i]);
                                WordAutomationObj.GotoastLinePoint();
                                WordAutomationObj.NextLine();
                                WordAutomationObj.GotoastLinePoint();
                            }

                        }

                    }

                }

            }

            //WordAutomationObj.final_Merge();

            // Overview_Page(WordAutomationObj);
            WordAutomationObj.SaveAs(@"D:\images\test11.doc");
            WordAutomationObj.oWord.Quit();
        }
        //public static void FirstPage(DataSet ds)
        //{
        //    WordAutomation WordAutomationObj = new WordAutomation();

        //    if (ds != null)
        //    {
        //        if (ds.Tables.Count > 0)
        //        {
        //            for (int i = 0; i < ds.Tables.Count; i++)
        //            {
        //                if (ds.Tables[i] != null)
        //                {
        //                    if (ds.Tables[i].Rows.Count > 0)
        //                    {
        //                        WordAutomationObj.GotoastLinePoint();
        //                        WordAutomationObj.CreateSimpleTable_1(ds.Tables[i].Rows.Count, 3, ds.Tables[i]);

        //                    }

        //                }

        //            }

        //        }

        //    }

        //    ////WordAutomationObj.final_Merge();

        //    //// Overview_Page(WordAutomationObj);
        //    //WordAutomationObj.SaveAs(@"D:\images\test12.doc");
        //    //WordAutomationObj.oWord.Quit();
        //}

        public static void OverViewPage(DataSet dsOverViewPage, WordAutomation WordAutomationObj)
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

                                    WordAutomationObj.PageHeading_New("Brand Overview", "");
                                   
                                    WordAutomationObj.InserLineBreak();

                                    string objBookmark = "B_OverViewPage";
                                   

                                    ErrorLog("Inside CreateSimpleTable_1", "Started");

                                    WordAutomationObj.CreateSimpleTable_1(dsOverViewPage.Tables[i].Rows.Count, 3, dsOverViewPage.Tables[i], objBookmark);

                                    ErrorLog("Inside CreateSimpleTable_1", "End");

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
        public static void PrintGraph(int CDID, WordAutomation WordAutomationObj)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/DossierChart.aspx with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/DossierChart.aspx?CDID=" + CDID+"&count=";
              
                
                for (int i = 1; i <= 4; i++)
                {
                    if (i == 1)
                    {

                        WordAutomationObj.PageHeading_New("Distribution by date", "");
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        WordAutomationObj.SaveImageFromUrl(URL + 6, (CDID + "_" + 6).ToString());
                    }
                    if (i == 2)
                    {

                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        WordAutomationObj.PageHeading_New("Distribution by publication type", "");
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        WordAutomationObj.SaveImageFromUrl(URL + i, (CDID + "_" + i).ToString());


                    }
                    if (i == 3)
                    {
                        WordAutomationObj.InsertPagebreak();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        WordAutomationObj.PageHeading_New("Distribution by language", "");
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        WordAutomationObj.SaveImageFromUrl(URL + i, (CDID + "_" + i).ToString());
                    }
                    if (i == 4)
                    {
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        WordAutomationObj.PageHeading_New("Distribution by sentiment", "");
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                        //WordAutomationObj.GotoastLinePoint();
                        //WordAutomationObj.InserLineBreak();
                        WordAutomationObj.SaveImageFromUrl_4_Graph(URL + i, (CDID + "_" + i).ToString());
                    }

                    

                }
                
                //int i = 5;
                //    if (i == 5)
                //    {
                //        WordAutomationObj.GotoastLinePoint();
                //        WordAutomationObj.InserLineBreak();
                //        WordAutomationObj.GotoastLinePoint();
                //        WordAutomationObj.InserLineBreak();
                //        WordAutomationObj.PageHeading_New("Sentiment", "");
                //        WordAutomationObj.GotoastLinePoint();
                //        WordAutomationObj.InserLineBreak();
                //        WordAutomationObj.GotoastLinePoint();
                //        WordAutomationObj.InserLineBreak();
                //    }

                //    WordAutomationObj.SaveImageFromUrl(URL+i, (CDID + "_" + i).ToString());

               

               

                 //WordAutomationObj.GotoastLinePoint();
                 //WordAutomationObj.InserLineBreak();          
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static void CoverageCategory_Graph (int CDID, WordAutomation WordAutomationObj)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/DossierChart.aspx with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/CoverageCategory_Graph.aspx?CDID=" + CDID + "&count=";

              //  for (int i = 1; i <= 4; i++)
               // {

                    WordAutomationObj.SaveImageFromUrl_Graph(URL + 1, (CDID + "_" + 1).ToString());
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();
              //  }


                          
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void ArticleReach_Graph(int CDID, WordAutomation WordAutomationObj)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/DossierChart.aspx with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/Template2BarGraph.aspx?CDID=" + CDID + "&count=";

                for (int i = 1; i <= 2; i++)
                {
                    if (i == 1)
                    {
                        WordAutomationObj.PageHeading_New("Circulation/Article Count-Print Media ", "");
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                        WordAutomationObj.SaveImageFromUrl_Graph(URL + i, (CDID + "_" + i).ToString());
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                    }
                    if (i == 2)
                    {
                        WordAutomationObj.PageHeading_New("Circulation/Article Count-Online Media", "");
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                        WordAutomationObj.SaveImageFromUrl_Graph(URL + i, (CDID + "_" + i).ToString());
                        WordAutomationObj.GotoastLinePoint();
                        WordAutomationObj.InserLineBreak();
                    }
                }


                          
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static void Article_TYPE_Graph(int CDID, WordAutomation WordAutomationObj)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/DossierChart.aspx with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/ArticleTYPE_Graph.aspx?CDID=" + CDID + "&count=";

               // for (int i = 1; i <= 6; i++)
               // {

                    WordAutomationObj.SaveImageFromUrl_Graph(URL + 1, (CDID + "_" + 1).ToString());
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();
               // }



            }
            catch (Exception)
            {

                throw;
            }
        }
        public static void Mention_TYPE_Graph(int CDID, WordAutomation WordAutomationObj)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/Template2HrGraph.aspx?CDID=226&count=1 with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/Template2HrGraph.aspx?CDID=" + CDID + "&count=";

                for (int i = 1; i <= 1; i++)
                {

                    WordAutomationObj.SaveImageFromUrl_Graph(URL + i, (CDID + "_" + i).ToString());
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();
                }



            }
            catch (Exception)
            {

                throw;
            }
        }
        public static void PublicationType_Graph(int CDID, WordAutomation WordAutomationObj)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/Template2HrGraph.aspx?CDID=226&count=1 with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/Template2HrGraph.aspx?CDID=" + CDID + "&count=";

               // for (int i = 2; i <= 2; i++)
                //{

                    WordAutomationObj.SaveImageFromUrl_Graph(URL + 1, (CDID + "_" + 1).ToString());
                    WordAutomationObj.GotoastLinePoint();
                    WordAutomationObj.InserLineBreak();
               // }



            }
            catch (Exception)
            {

                throw;
            }
        }
        public static void PublicationType_Graph1(int CDID, WordAutomation WordAutomationObj)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/Template2HrGraph.aspx?CDID=226&count=1 with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/PublicationType_Graph.aspx?CDID=" + CDID + "&count=";

                // for (int i = 2; i <= 2; i++)
                //{

                WordAutomationObj.SaveImageFromUrl_Graph(URL + 1, (CDID + "_" + 1).ToString());
                WordAutomationObj.GotoastLinePoint();
                WordAutomationObj.InserLineBreak();
                // }



            }
            catch (Exception)
            {

                throw;
            }
        }
        //public static void Overview_Page(WordAutomation WordAutomationObj)
        //{
        //    WordAutomationObj.CreateSimpleTable(WordAutomationObj.oDoc, 2, 3);

        //    WordAutomationObj.oDoc.Paragraphs.Add();

        //    // Add colspan-like effect in the 4th row
        //    WordAutomationObj.AddColspanEffect(WordAutomationObj.oDoc, 2);        
        //}


        #region Log

        private static void ErrorLog(string functionName, string error)
        {
            try
            {
                string LogPath = ConfigurationManager.AppSettings["ErrorLogPath"];

               
                using (TextWriter myWriter = File.AppendText(LogPath))
                {

                    TextWriter.Synchronized(myWriter).WriteLine("<--> :{0}", string.Empty);
                    TextWriter.Synchronized(myWriter).WriteLine("Function Name :{0}", functionName);
                    TextWriter.Synchronized(myWriter).WriteLine("Error :{0}", error);
                    TextWriter.Synchronized(myWriter).WriteLine("Date :{0}", DateTime.Now.ToString());
                    TextWriter.Synchronized(myWriter).WriteLine("<--> :{0}", string.Empty);
                    TextWriter.Synchronized(myWriter).Flush();
                    TextWriter.Synchronized(myWriter).Close();
                }
            }
            catch (Exception ex)
            {
            }
        }

        #endregion

        private static void TimeLog(string strtext)
        {
            try
            {
                string LogPath = ConfigurationManager.AppSettings["TimeLogPath"];
                using (TextWriter myWriter = File.AppendText(LogPath))
                {

                    TextWriter.Synchronized(myWriter).WriteLine("<--> :{0}", string.Empty);
                    TextWriter.Synchronized(myWriter).WriteLine("TimeLog :{0}", strtext);
                    TextWriter.Synchronized(myWriter).WriteLine("Date :{0}", DateTime.Now.ToString());
                    TextWriter.Synchronized(myWriter).WriteLine("<--> :{0}", string.Empty);
                    TextWriter.Synchronized(myWriter).Flush();
                    TextWriter.Synchronized(myWriter).Close();
                }
            }
            catch (Exception ex)
            {
                ErrorLog("TimeLog", ex.ToString());
            }
        }

        private static bool SendMail(string ToMailID, string Body, string Subject, string FilePath, string CCMailID, string BCCMailID)
        {

            TimeLog("ToMailID " + ToMailID);

            #region Log
            //TimeLog("Body " + Body);
            TimeLog("Subject " + Subject);
            TimeLog("CCMailID " + CCMailID);
            TimeLog("BCCMailID " + BCCMailID);
            #endregion

            bool IsSent = false;
            try
            {
                //Get From App Config From MailID
                string FromMailID = ConfigurationManager.AppSettings["FromMailAddress"].ToString();
                //Set From Mail ID To MailAddress Properties
                MailAddress ObjFrom = new MailAddress(FromMailID, ConfigurationManager.AppSettings["FromDisplayName"].ToString());
                //Call the refrence of ddll system.net.mail.MailMessage which help to send you mail
                System.Net.Mail.MailMessage mails = new System.Net.Mail.MailMessage();
                //Define all proprties related mail configration
                //Set from mail ID here
                mails.From = ObjFrom;
                //Multiple TO Mailid which available comma sepreated the split bu (,)
                string[] SplitTOMailID = ToMailID.Split(',');
                //if array is available the using foreach loop get one by one id 
                foreach (string ToMailIDValues in SplitTOMailID)
                {
                    //Check here if id is not null or blank
                    if (!string.IsNullOrEmpty(ToMailIDValues))
                    {
                        //Here add in To MailID Properites
                        mails.To.Add(ToMailIDValues);
                    }
                }
                //Multiple CC Mailid which available comma sepreated the split bu (,)
                string[] SplitCCMailID = CCMailID.Split(',');
                foreach (string CCMailIDValues in SplitCCMailID)
                {
                    //Check here if id is not null or blank
                    if (!string.IsNullOrEmpty(CCMailIDValues))
                    {
                        //Here add in CC MailID Properites
                        mails.CC.Add(CCMailIDValues);
                    }
                }
                //Multiple BCC Mailid which available comma sepreated the split bu (,)
                string[] SplitBCCMailID = BCCMailID.Split(',');
                foreach (string CCMailIDValues in SplitBCCMailID)
                {
                    //Check here if id is not null or blank
                    if (!string.IsNullOrEmpty(CCMailIDValues))
                    {
                        //Here add in BCC MailID Properites
                        mails.Bcc.Add(CCMailIDValues);
                    }
                }
                // Set Subject
                mails.Subject = Subject.Replace('\r', ' ').Replace('\n', ' ');
                mails.Subject = HttpUtility.HtmlDecode(mails.Subject);
                //if Set html body then always set true
                mails.IsBodyHtml = true;
                //set Message Body
                mails.Body = Body;

                #region Attachment
                //attachment file is exist or not check
                if (File.Exists(FilePath))
                {
                    //if yes the attach the excel sheet
                    Attachment attach = new Attachment(FilePath);
                    //add in mail
                    mails.Attachments.Add(attach);
                }
                #endregion


                //Call SMTP as alias objSMTP
                SmtpClient objSMTP = new SmtpClient();

                //SET SMTP Host from mail Config
                objSMTP.Host = ConfigurationManager.AppSettings["MailServer"].ToString();
                //SET SMTP Port From Mail Config
                objSMTP.Port = Convert.ToInt32(ConfigurationManager.AppSettings["MailPort"]);
                //SET SMTP Userdefault credentails
                objSMTP.UseDefaultCredentials = true;


                TimeLog("Sending Mail To Server");
                // and finally send mail
                objSMTP.Send(mails);
                TimeLog("Mail Sent");
                IsSent = true;
            }
            //catch (SmtpException ex)
            //{
            //    ErrorLog("SendMail ", ex.ToString() + " ->> " + ex.Message + " ->> " + ex.StackTrace);
            //    //SendMail(ConfigurationManager.AppSettings["MailErrorID"].ToString(), ex.ToString(), "Error In SendMail Function", "", "", "");
            //}
            catch (Exception ex)
            {
                ErrorLog("SendMail ", ex.ToString() + " ->> " + ex.Message + " ->> " + ex.StackTrace);

                //throw err;
            }

            return IsSent;
        }

        private static bool ErrorSendMail(string ToMailID, string Body, string Subject, string FilePath, string CCMailID, string BCCMailID)
        {

            TimeLog("ToMailID " + ToMailID);

            #region Log
            //TimeLog("Body " + Body);
            TimeLog("Subject " + Subject);
            TimeLog("CCMailID " + CCMailID);
            TimeLog("BCCMailID " + BCCMailID);
            #endregion

            bool IsSent = false;
            try
            {
                //Get From App Config From MailID
                string FromMailID = ConfigurationManager.AppSettings["FromMailAddress"].ToString();
                //Set From Mail ID To MailAddress Properties
                MailAddress ObjFrom = new MailAddress(FromMailID, ConfigurationManager.AppSettings["FromDisplayName"].ToString());
                //Call the refrence of ddll system.net.mail.MailMessage which help to send you mail
                System.Net.Mail.MailMessage mails = new System.Net.Mail.MailMessage();
                //Define all proprties related mail configration
                //Set from mail ID here
                mails.From = ObjFrom;
                //Multiple TO Mailid which available comma sepreated the split bu (,)
                string[] SplitTOMailID = ToMailID.Split(',');
                //if array is available the using foreach loop get one by one id 
                foreach (string ToMailIDValues in SplitTOMailID)
                {
                    //Check here if id is not null or blank
                    if (!string.IsNullOrEmpty(ToMailIDValues))
                    {
                        //Here add in To MailID Properites
                        mails.To.Add(ToMailIDValues);
                    }
                }
                //Multiple CC Mailid which available comma sepreated the split bu (,)
                string[] SplitCCMailID = CCMailID.Split(',');
                foreach (string CCMailIDValues in SplitCCMailID)
                {
                    //Check here if id is not null or blank
                    if (!string.IsNullOrEmpty(CCMailIDValues))
                    {
                        //Here add in CC MailID Properites
                        mails.CC.Add(CCMailIDValues);
                    }
                }
                //Multiple BCC Mailid which available comma sepreated the split bu (,)
                string[] SplitBCCMailID = BCCMailID.Split(',');
                foreach (string CCMailIDValues in SplitBCCMailID)
                {
                    //Check here if id is not null or blank
                    if (!string.IsNullOrEmpty(CCMailIDValues))
                    {
                        //Here add in BCC MailID Properites
                        mails.Bcc.Add(CCMailIDValues);
                    }
                }
                // Set Subject
                mails.Subject = Subject.Replace('\r', ' ').Replace('\n', ' ');
                mails.Subject = HttpUtility.HtmlDecode(mails.Subject);
                //if Set html body then always set true
                mails.IsBodyHtml = true;
                //set Message Body
                mails.Body = Body;

                #region Attachment
                //attachment file is exist or not check
                if (File.Exists(FilePath))
                {
                    //if yes the attach the excel sheet
                    Attachment attach = new Attachment(FilePath);
                    //add in mail
                    mails.Attachments.Add(attach);
                }
                #endregion


                //Call SMTP as alias objSMTP
                SmtpClient objSMTP = new SmtpClient();

                //SET SMTP Host from mail Config
                objSMTP.Host = ConfigurationManager.AppSettings["MailServer"].ToString();
                //SET SMTP Port From Mail Config
                objSMTP.Port = Convert.ToInt32(ConfigurationManager.AppSettings["MailPort"]);
                //SET SMTP Userdefault credentails
                objSMTP.UseDefaultCredentials = true;


                TimeLog("Sending Mail To Server");
                // and finally send mail
                objSMTP.Send(mails);
                TimeLog("Mail Sent");
                IsSent = true;
            }
            //catch (SmtpException ex)
            //{
            //    ErrorLog("SendMail ", ex.ToString() + " ->> " + ex.Message + " ->> " + ex.StackTrace);
            //    //SendMail(ConfigurationManager.AppSettings["MailErrorID"].ToString(), ex.ToString(), "Error In SendMail Function", "", "", "");
            //}
            catch (Exception ex)
            {
                ErrorLog("SendMail ", ex.ToString() + " ->> " + ex.Message + " ->> " + ex.StackTrace);

                //throw err;
            }

            return IsSent;
        }


        #region Images
        public static List<NewsArticle_Online> Process_Images(List<NewsArticle_Online> lstArticles,Int32 CDID)
        {
            List<NewsArticle_Online> lstFinal = new List<NewsArticle_Online>();

            try
            {

                int counter = 0;
                var stopWatch = Stopwatch.StartNew();
                ParallelOptions po = new ParallelOptions();
                po.MaxDegreeOfParallelism = Convert.ToInt32(ConfigurationManager.AppSettings["Image_MaxDegreeOfParallelism"].ToString());
                System.Net.ServicePointManager.DefaultConnectionLimit = Convert.ToInt32(ConfigurationManager.AppSettings["Image_MaxDegreeOfParallelism"].ToString());

                Parallel.For(0, lstArticles.Count, po, i =>
                {
                    try
                    {
                        TimeLog("Start CreateImage");
                        lstArticles[i].screenshot_fileName = CreateImage(lstArticles[i].NewsURL, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                        lstFinal.Add(lstArticles[i]);
                        TimeLog("End CreateImage");
                        counter++;
                    }
                    catch (Exception ex)
                    {
                        ErrorLog("Parallel Error for " + Convert.ToString(lstArticles[i].NewsURL.ToString()), ex.ToString());
                    }

                });
                TimeLog("End Parallel.For");
                TimeLog("Parallel.For() execution time for " + " = {0} seconds:" + stopWatch.Elapsed.TotalSeconds);

            }
            catch (Exception ex)
            {
                ErrorLog("Process_Images ", ex.Message + " ->> " + ex.StackTrace);
            }

            return lstFinal;
        }

        private static string CreateImage(string strUrl, string strFileName)
        {
            string strFileNameReturn = string.Empty;
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                //string strFetchurl = "https://api.urlbox.io/v1/4G9b0Mqi6yniUmhs/png?url=" + HttpUtility.UrlEncode(strUrl) + "&delay=6000";
                string strFetchurl = System.Configuration.ConfigurationManager.AppSettings["Screenshot_APIUrl"].Replace("&amp;", "&").Replace("myurl", HttpUtility.UrlEncode(strUrl));
                strFetchurl = strFetchurl.Replace("\r\n", "");

                TimeLog(strFetchurl);

                WebRequest request = WebRequest.Create(strFetchurl);

                //// Get the response. 
                WebResponse response = request.GetResponse();


                // Get the stream containing content returned by the server.
                Stream dataStream = response.GetResponseStream();
                System.Drawing.Image image = System.Drawing.Image.FromStream(dataStream);
                string strPath = System.Configuration.ConfigurationManager.AppSettings["ArticleOnlineIamgeURL"] + strFileName + ".jpg";//ConfigurationManager.AppSettings["ScreenshotExtension"];
                image.Save(strPath);
                strFileNameReturn = strFileName + ".jpg";

                
            }
            catch (Exception ex)
            {
                ErrorLog("CreateImage", strUrl + " ->" + ex.ToString());
            }

            return strFileNameReturn;
        }

        public static void  Delete_All_File_Images_In_Folder()
        {

         DirectoryInfo dir = new DirectoryInfo(ConfigurationManager.AppSettings["ArticleOnlineIamgeURL"]);

                 foreach (FileInfo fi in dir.GetFiles("*.jpg"))
                        {

                            fi.Delete();
                        }
        }
        #endregion


    }

}
