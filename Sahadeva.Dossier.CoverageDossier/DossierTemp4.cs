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
    class DossierTemp4
    {
        public void MainDT4(DataTable lstCDID)
        {
            string ClientEmailIdsTo = "";
            string EntityName = "";
            string EmailIdsCC = "";
            string EmailIdsBCC = "";

            try
            {

                Dossier.DAL.DossierTemp1DAL dalDT1 = new DAL.DossierTemp1DAL();
                //lstCDID.Rows.Add(7803);
                if (lstCDID.Rows.Count > 0)
                {
                    string[] sep = new string[] { "\r\n" };
                    var stopWatch = Stopwatch.StartNew();
                    ParallelOptions po = new ParallelOptions();
                    po.MaxDegreeOfParallelism = Convert.ToInt32(ConfigurationManager.AppSettings["MaxDegreeOfParallelism"]);
                    System.Net.ServicePointManager.DefaultConnectionLimit = Convert.ToInt32(ConfigurationManager.AppSettings["MaxDegreeOfParallelism"]); Parallel.For(0, lstCDID.Rows.Count, po, i =>
                    {
                        try
                        {
                            List<NewsArticle> lstNewsArticle = new List<NewsArticle>();
                            List<NewsArticle_Online> lstNewsArticleOnline = new List<NewsArticle_Online>();

                            #region MyRegion
                            ErrorLog("Dossier_FetchClientData", "Started");
                            DataTable dtClient = dalDT1.Dossier_FetchClientData_DT1(Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                            ErrorLog("Dossier_FetchClientData", "End");

                            ErrorLog("Dossier_FetchArticleData_Print", "Started");
                            lstNewsArticle = dalDT1.Dossier_FetchArticleData_Print_DT1(Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                            ErrorLog("Dossier_FetchArticleData_Print", "End");

                            ErrorLog("Dossier_FetchArticleData_Online", "Started");
                            lstNewsArticleOnline = dalDT1.Dossier_FetchArticleData_Online_DT1(Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                            ErrorLog("Dossier_FetchArticleData_Online", "End");

                            if (lstNewsArticle.Count > 0 || lstNewsArticleOnline.Count > 0)
                            {
                                ErrorLog("CreateWord", "Started");
                                dalDT1.CoverageDossier_UpdateStatus(Convert.ToInt32(lstCDID.Rows[i]["CDID"]), Convert.ToInt32(ConfigurationManager.AppSettings["DossierGenerationStart"]));


                                ErrorLog("CreateScreenShorts", "Started");
                                Process_Images(lstNewsArticleOnline, Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                                ErrorLog("CreateScreenShorts", "End");

                                CreateWord(lstNewsArticle, lstNewsArticleOnline, dtClient, Convert.ToInt32(lstCDID.Rows[i]["CDID"]));
                                ErrorLog("CreateWord", "End");

                                string MailSentStatusID = Convert.ToString(ConfigurationManager.AppSettings["MailSentStatusID"]);
                                ErrorLog("CoverageDossier_UpdateStatus", "Started");
                                dalDT1.CoverageDossier_UpdateStatus(Convert.ToInt32(lstCDID.Rows[i]["CDID"]), Convert.ToInt32(MailSentStatusID));
                                ErrorLog("CoverageDossier_UpdateStatus", "End");

                                ErrorLog("Delete all File in folder", "Started");
                                Delete_All_File_Images_In_Folder();

                                ErrorLog("Delete all File in folder", "End");
                            }
                            else
                            {
                                ErrorLog("Mail Sent", "No Article Data Found");
                                SendMail(Convert.ToString(dtClient.Rows[0]["EmailIdsTo"]), "No Articles Data Found", "Dossier Generator - " + Convert.ToString(dtClient.Rows[0]["EntityName"]), "", Convert.ToString(dtClient.Rows[0]["EmailIdsCC"]), Convert.ToString(dtClient.Rows[0]["EmailIdsBCC"]));
                                string MailSentStatusID = Convert.ToString(ConfigurationManager.AppSettings["MailSentStatusID"]);
                                dalDT1.CoverageDossier_UpdateStatus(Convert.ToInt32(lstCDID.Rows[i]["CDID"]), Convert.ToInt32(MailSentStatusID));
                                ErrorLog("Mail Sent", "No Article Data Found");
                            }

                            ClientEmailIdsTo = Convert.ToString(dtClient.Rows[0]["EmailIdsTo"]);
                            EntityName = Convert.ToString(dtClient.Rows[0]["EntityName"]);
                            EmailIdsCC = Convert.ToString(dtClient.Rows[0]["EmailIdsCC"]);
                            EmailIdsBCC = Convert.ToString(dtClient.Rows[0]["EmailIdsBCC"]);
                            #endregion
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
                    ErrorSendMail(ClientEmailIdsTo, "No Articles Data Found", "Dossier Generator - " + EntityName, "", EmailIdsCC, EmailIdsBCC);
                }
                ErrorSendMail(Convert.ToString(ConfigurationManager.AppSettings["ErrorEmailID"]), ex.ToString(), "Error Log Dossier Generator - " + EntityName, "", EmailIdsCC, EmailIdsBCC);
                
            }
            
        }

        public static void CreateWord(List<NewsArticle> lstNewsArticle, List<NewsArticle_Online> lstNewsArticleOnline, DataTable dtClient, Int32 CDID)
        {
            try
            {

                WordAutomationTemp4 WordAutomationObjTemp4 = new WordAutomationTemp4();

                string BookMark = string.Empty;

                Dossier.DAL.DossierTemp1DAL dalDT1 = new DAL.DossierTemp1DAL();


                WordAutomationObjTemp4.GotoastLinePoint();
                string FirstpagebgURL = ConfigurationManager.AppSettings["Firstpagebg"];
               
                WordAutomationObjTemp4.FirstPage_New(Convert.ToString(dtClient.Rows[0]["EntityName"]), Convert.ToString(dtClient.Rows[0]["Description"]), Convert.ToString(dtClient.Rows[0]["FromDate"]) + " to " + Convert.ToString(dtClient.Rows[0]["ToDate"]), FirstpagebgURL);
                WordAutomationObjTemp4.InsertPagebreak();


                DataSet dsOverViewPage = dalDT1.USP_CoverageDossier_OverView_Page_Data_DT4(CDID);

                OverViewPage(dsOverViewPage, WordAutomationObjTemp4);            //Overview Page
                
                WordAutomationObjTemp4.PageHeading_New("Overview summary", "");
                WordAutomationObjTemp4.GotoastLinePoint();
                WordAutomationObjTemp4.InserLineBreak();

                WordAutomationObjTemp4.InsertText(Convert.ToString(dtClient.Rows[0]["Summary"]));

                WordAutomationObjTemp4.GotoastLinePoint();
                WordAutomationObjTemp4.InserLineBreak();
                WordAutomationObjTemp4.InsertPagebreak();

                //PrintGraph(CDID, WordAutomationObjTemp4);        // Overview Print Graph 

                WordAutomationObjTemp4.GotoastLinePoint();
                WordAutomationObjTemp4.InsertPagebreak();

                ErrorLog("Inside CreateWord - Print", "Started");

                #region For Print

                WordAutomationObjTemp4.GotoastLinePoint();
                WordAutomationObjTemp4.InserLineBreak();

                WordAutomationObjTemp4.PageHeading_New("Table of contents - print", "B_PageHeadingTable_Print");
                WordAutomationObjTemp4.GotoastLinePoint();
                WordAutomationObjTemp4.InserLineBreak();

                if (lstNewsArticle.Count != 0)
                {

                    WordAutomationObjTemp4.InsertTable_Articles_Print(lstNewsArticle);
                    WordAutomationObjTemp4.GotoastLinePoint();
                    WordAutomationObjTemp4.InserLineBreak();
                    WordAutomationObjTemp4.InsertPagebreak();
                }

                WordAutomationObjTemp4.PageHeading_New("Prominent Coverages - Print", "B_PageHeadingCoverages_Print");
                WordAutomationObjTemp4.GotoastLinePoint();
                WordAutomationObjTemp4.InserLineBreak();

                for (int i = 0; i < lstNewsArticle.Count; i++)
                {
                    NewsArticle objNewsArticle = new NewsArticle();

                    objNewsArticle = lstNewsArticle[i];

                    WordAutomationObjTemp4.InsertHeaderTable_New_Print(objNewsArticle);

                    WordAutomationObjTemp4.GotoastLinePoint();
                    WordAutomationObjTemp4.InserLineBreak();

                    WordAutomationObjTemp4.InsertBoldTextLable("Article synopsis - ");
                    WordAutomationObjTemp4.InsertText(Convert.ToString(objNewsArticle.NewsText));
                    WordAutomationObjTemp4.GotoastLinePoint();
                
                    WordAutomationObjTemp4.GotoastLinePoint();
                    WordAutomationObjTemp4.InserLineBreak();

                    //string BookMark = "B" + 1;
                    string ArticleIamgeURL = ConfigurationManager.AppSettings["ArticleIamgeURL"];

                    if (objNewsArticle.NewsURL != null)
                    {
                        BookMark = "B_Print" + lstNewsArticle.IndexOf(lstNewsArticle[i]).ToString();

                        WordAutomationObjTemp4.InsertArticleImageIntoBody(ArticleIamgeURL + objNewsArticle.IMG_URL, BookMark);
                        WordAutomationObjTemp4.GotoastLinePoint();
                    }
                    
                    WordAutomationObjTemp4.InsertPagebreak();

                }

                string EntityID = Convert.ToString(dtClient.Rows[0]["EntityID"]);
                string ClientLogoURL = Convert.ToString(ConfigurationManager.AppSettings["ClientLogoURL"]);
                string ADFLogoURL = Convert.ToString(ConfigurationManager.AppSettings["ADFLogoURL"]);

                //Logo check 
                if (EntityID != null || EntityID != string.Empty)
                {
                    if ((System.IO.File.Exists(ClientLogoURL + EntityID + ".jpg")))
                    {
                        WordAutomationObjTemp4.InsertHeaderandFooter(ClientLogoURL + EntityID + ".jpg", ADFLogoURL);
                    }
                    else
                    {
                        WordAutomationObjTemp4.InsertHeaderandFooter(ADFLogoURL, ADFLogoURL);
                    }
                }

                #endregion

                ErrorLog("Inside CreateWord - Print", "End");
                ErrorLog("Inside CreateWord - Online", "Started");

                #region For Online
                if (lstNewsArticleOnline.Count != 0)
                {


                    WordAutomationObjTemp4.GotoastLinePoint();
                    WordAutomationObjTemp4.InserLineBreak();

                    WordAutomationObjTemp4.PageHeading_New("Table of contents - online", "B_PageHeadingTable_Online");
                    WordAutomationObjTemp4.GotoastLinePoint();
                    WordAutomationObjTemp4.InserLineBreak();

                    if (lstNewsArticleOnline.Count != 0)
                    {
                        WordAutomationObjTemp4.InsertTable_Articles_Online(lstNewsArticleOnline);
                        WordAutomationObjTemp4.GotoastLinePoint();
                        WordAutomationObjTemp4.InserLineBreak();
                        WordAutomationObjTemp4.InsertPagebreak();
                    }


                    WordAutomationObjTemp4.PageHeading_New("Prominent Coverages - Online", "B_PageHeadingCoverages_Online");
                  
                    WordAutomationObjTemp4.GotoastLinePoint();
                    WordAutomationObjTemp4.InserLineBreak();

                    for (int i = 0; i < lstNewsArticleOnline.Count; i++)
                    {
                        NewsArticle_Online objNewsArticle = new NewsArticle_Online();

                        objNewsArticle = lstNewsArticleOnline[i];

                        WordAutomationObjTemp4.GotoastLinePoint();
                        WordAutomationObjTemp4.InserLineBreak();
                                                
                        WordAutomationObjTemp4.InsertHeaderTable_New_Online(objNewsArticle);

                        WordAutomationObjTemp4.GotoastLinePoint();

                        WordAutomationObjTemp4.InserLineBreak();
                     
                        WordAutomationObjTemp4.InsertBoldTextLable("Article synopsis - ");

                        WordAutomationObjTemp4.InsertText(Convert.ToString(objNewsArticle.NewsText));
                        WordAutomationObjTemp4.GotoastLinePoint();
                      
                        WordAutomationObjTemp4.GotoastLinePoint();
                       

                        string ArticleIamgeURL = ConfigurationManager.AppSettings["ArticleOnlineIamgeURL"];

                        if (objNewsArticle.screenshot_fileName != null)
                        {
                            BookMark = "B_Online" + lstNewsArticleOnline.IndexOf(lstNewsArticleOnline[i]).ToString();
                            string ext = Path.GetExtension(ArticleIamgeURL + objNewsArticle.screenshot_fileName);
                            if (ext != null)
                            {
                              
                                WordAutomationObjTemp4.InsertArticleImageIntoBody(ArticleIamgeURL + objNewsArticle.screenshot_fileName, BookMark);
                            }
                            else
                            {
                                WordAutomationObjTemp4.InsertArticleImageIntoBody(ArticleIamgeURL + objNewsArticle.screenshot_fileName + ".jpg", BookMark);
                            }
                         
                            WordAutomationObjTemp4.GotoastLinePoint();
                        }
                      
                        WordAutomationObjTemp4.InsertPagebreak();
                    
                    }
                }              
                #endregion

                ErrorLog("Inside CreateWord - Online", "End");

                String DossierFileName = Convert.ToString(dtClient.Rows[0]["EntityName"]).Replace(" ", "_").Trim() + "_" +
                                 DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString()
                                + "_" + Convert.ToString(dtClient.Rows[0]["FromDate"]).Replace(" ", "_") +
                                "_" + Convert.ToString(dtClient.Rows[0]["ToDate"]).Replace(" ", "_")
                                + Convert.ToString(dtClient.Rows[0]["CDID"]);

               
                string SavePath = Convert.ToString(ConfigurationManager.AppSettings["SavePath"]);
                WordAutomationObjTemp4.SaveAs(SavePath + DossierFileName + ".doc");
                WordAutomationObjTemp4.oWord.Quit();
                string OutputFiles_URL = Convert.ToString(ConfigurationManager.AppSettings["OutputFiles_URL_DT4"]);
               
                string EmailBody = System.IO.File.ReadAllText(ConfigurationManager.AppSettings["MailBodypath"].ToString());

                EmailBody = EmailBody.Replace("#DossierNo#", "ED-" + Convert.ToString(dtClient.Rows[0]["CDCID"]));
                EmailBody = EmailBody.Replace("#link#", OutputFiles_URL + DossierFileName + ".doc");
                EmailBody = EmailBody.Replace("#EntityName#", Convert.ToString(dtClient.Rows[0]["EntityName"]));
                EmailBody = EmailBody.Replace("#EmpName#", Convert.ToString(dtClient.Rows[0]["Name"]));


                dalDT1.Update_CD_OutputFileLink(Convert.ToInt32(dtClient.Rows[0]["CDID"]), OutputFiles_URL + DossierFileName + ".doc");

               // SendMail(Convert.ToString(dtClient.Rows[0]["EmailIdsTo"]), EmailBody, "Dossier Generator - " + Convert.ToString(dtClient.Rows[0]["EntityName"]), "", Convert.ToString(dtClient.Rows[0]["EmailIdsCC"]), Convert.ToString(dtClient.Rows[0]["EmailIdsBCC"] + "," + ConfigurationManager.AppSettings["ErrorEmailID"]));
            }
            catch (Exception)
            {
                throw;
            }
        }


        public static void Test(DataSet ds)
        {
            WordAutomationTemp4 WordAutomationObj = new WordAutomationTemp4();

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
                                WordAutomationObj.CreateSimpleTable(ds.Tables[i].Rows.Count, 3, ds.Tables[i]);
                                WordAutomationObj.GotoastLinePoint();
                                WordAutomationObj.NextLine();
                                WordAutomationObj.GotoastLinePoint();
                            }

                        }

                    }

                }

            }
           
            WordAutomationObj.SaveAs(@"D:\images\test11.doc");
            WordAutomationObj.oWord.Quit();
        }


        public static void OverViewPage(DataSet dsOverViewPage, WordAutomationTemp4 WordAutomationObjTemp4)
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

                                    WordAutomationObjTemp4.PageHeading_New("Impact Table", "");

                                    WordAutomationObjTemp4.InserLineBreak();

                                    string objBookmark = "B_OverViewPage";


                                    ErrorLog("Inside CreateSimpleTable_1", "Started");

                                    WordAutomationObjTemp4.CreateSimpleTable_1(dsOverViewPage.Tables[i].Rows.Count, 3, dsOverViewPage.Tables[i], objBookmark);

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
        public static void PrintGraph(int CDID, WordAutomationTemp4 WordAutomationObjTemp4)
        {
            try
            {
                ////Call https://aidossier.adfactorspr.com/DossierChart.aspx with GraphCount, CDID 
                string URL = "https://aidossier.adfactorspr.com/DossierChart.aspx?CDID=" + CDID + "&count=";

                for (int i = 1; i <= 4; i++)
                {
                    if (i == 1)
                    {

                        WordAutomationObjTemp4.PageHeading_New("Distribution by date", "");
                        WordAutomationObjTemp4.GotoastLinePoint();
                        WordAutomationObjTemp4.InserLineBreak();                       
                        WordAutomationObjTemp4.SaveImageFromUrl(URL + 6, (CDID + "_" + 6).ToString());
                    }
                    if (i == 2)
                    {

                        WordAutomationObjTemp4.PageHeading_New("Distribution by publication type", "");
                        WordAutomationObjTemp4.GotoastLinePoint();
                        WordAutomationObjTemp4.InserLineBreak();                       
                        WordAutomationObjTemp4.SaveImageFromUrl(URL + i, (CDID + "_" + i).ToString());


                    }
                    if (i == 3)
                    {
                        WordAutomationObjTemp4.InsertPagebreak();
                        WordAutomationObjTemp4.GotoastLinePoint();
                        WordAutomationObjTemp4.InserLineBreak();
                        WordAutomationObjTemp4.PageHeading_New("Distribution by language", "");
                        WordAutomationObjTemp4.GotoastLinePoint();
                        WordAutomationObjTemp4.InserLineBreak();
                      
                        WordAutomationObjTemp4.SaveImageFromUrl(URL + i, (CDID + "_" + i).ToString());
                    }
                    if (i == 4)
                    {
                        WordAutomationObjTemp4.PageHeading_New("Distribution by sentiment", "");
                        WordAutomationObjTemp4.GotoastLinePoint();
                        WordAutomationObjTemp4.InserLineBreak();                        
                        WordAutomationObjTemp4.SaveImageFromUrl_4_Graph(URL + i, (CDID + "_" + i).ToString());
                    }
                }

            }
            catch (Exception)
            {

                throw;
            }
        }
      
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


        public static List<NewsArticle_Online> Process_Images(List<NewsArticle_Online> lstArticles, Int32 CDID)
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

        public static void Delete_All_File_Images_In_Folder()
        {

            DirectoryInfo dir = new DirectoryInfo(ConfigurationManager.AppSettings["ArticleOnlineIamgeURL"]);

            foreach (FileInfo fi in dir.GetFiles("*.jpg"))
            {

                fi.Delete();
            }
        }
    }
}
