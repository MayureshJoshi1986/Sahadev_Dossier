using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sahadeva.Dossier.Common
{
    public class DatabaseConstants
    {

        #region General Constants
        public const string RadarMTSConnection = "RadarMTSConnection";
        public const string ACEConnectionString = "ACEConnectionString";
        public const string ConnectionString = "ConnectionString";


        #endregion

        #region Stored Procedures

        public const string USP_CoverageDossier_OverView_Page_Data_DT2 = "USP_CoverageDossier_OverView_Page_Data_DT2";
        public const string USP_CoverageDossier_OverView_Page_1 = "USP_CoverageDossier_OverView_Page_1";
        public const string USP_CoverageDossier_OverView_Page = "USP_CoverageDossier_OverView_Page";
        public const string USP_Dossier_FetchPrintArticleData = "USP_Dossier_FetchPrintArticleData";
        public const string USP_Fetch_DataByCDID_DT2 = "USP_Fetch_DataByCDID_DT2";
        public const string USP_FetchPending_DCIDsToProcess_DT2 = "USP_FetchPending_DCIDsToProcess_DT2";
        public const string USP_CoverageDossier_UpdateStatus = "USP_CoverageDossier_UpdateStatus";
        public const string USP_Insert_DossierGanrationLog = "USP_Insert_DossierGanrationLog";
        public const string USP_FetchPending_DCIDsToProcess_All = "USP_FetchPending_DCIDsToProcess_All";

        

        //Look Ups
        public const string LookUp_ID = "ID";
        public const string LookUp_DisplayText = "DisplayText";


        //OfflineNews
        public const string USP_OfflineNews_Update = "USP_OfflineNews_Update";
        public const string USP_OfflineNews_Insert = "USP_OfflineNews_Insert";



        //SPS CREATED WHILE IMPLEMENTING DOSSIER INTO NEW RADAR MTS (MTRACK)
        public const string USP_FetchClients = "USP_FetchClients";

        public const string USP_Update_CD_OutputFileLink = "USP_Update_CD_OutputFileLink";

        #endregion

        #region Stored Procedures temp1
        public const string USP_CoverageDossier_OverView_Page_Data_DT1 = "USP_CoverageDossier_OverView_Page_Data";
        //public const string USP_CoverageDossier_OverView_Page_1 = "USP_CoverageDossier_OverView_Page_1";
        //public const string USP_CoverageDossier_OverView_Page = "USP_CoverageDossier_OverView_Page";
        //public const string USP_Dossier_FetchPrintArticleData = "USP_Dossier_FetchPrintArticleData";
        public const string USP_Fetch_DataByCDID_DT1 = "USP_Fetch_DataByCDID";
        public const string USP_FetchPending_DCIDsToProcess_DT1 = "USP_FetchPending_DCIDsToProcess";
        //public const string USP_CoverageDossier_UpdateStatus = "USP_CoverageDossier_UpdateStatus";

        #endregion


        #region Stored Procedures temp1
        public const string USP_CoverageDossier_OverView_Page_Data_DT4 = "USP_CoverageDossier_OverView_Page_Data_DT4";


        #endregion

        #region Variables Name

        public const string CDID = "CDID";
        public const string NewsText = "NewsText";
        public const string NewsURL = "NewsURL";
        public const string Topics = "Topics";
        public const string Sentiment = "Sentiment";
        public const string Publication = "Publication";

        public const string Cluster = "Cluster";
        public const string Title = "Title";
        public const string NewsTraffic = "NewsTraffic";

        public const string DEID = "DEID";
        public const string Count = "Count";
        public const string Rowno = "Rowno";
        public const string Group_id = "Group_id";
        public const string CompanyID = "CompanyID";
        public const string FromDate = "FromDate";
        public const string Todate = "Todate";
        public const string GroupName = "group_name";
        public const string comp_name = "comp_name";//ClientID, ClientName
        public const string Comp_id = "Comp_id";
        //public const string comp_name = "ClientName";
        //public const string Comp_id = "ClientID";
        public const string CompNewsID = "CompNewsID";
        public const string News_id = "News_id";
        public const string News_dt = "News_dt";
        public const string page = "page";
        public const string NPos_ID = "NPos_ID";
        public const string edition = "edition";
        public const string NewsPosition = "NewsPosition";
        public const string newspaper = "newspaper";
        public const string News_Date = "News_dt";
        public const string MediaType = "NewsCategoryId";
        public const string headline = "headline";
        public const string ImageName = "ImageName";
        public const string Journalist = "Journalist";
        public const string summary = "summary";
        public const string NewsPaperid = "NewsPaper_id";
        public const string no_col = "no_col";
        public const string ht_cm = "ht_cm";
        public const string img_url = "img_url";
        public const string LanguageName = "LanguageName";
        public const string Lang_ID = "Lang_ID";
        //public const string LanguageName = "Language";
        //public const string Lang_ID = "LanguageID";
        public const string CompanyName = "CompanyName";
        public const string StartRowIndex = "StartRowIndex";
        public const string EndRowIndex = "EndRowIndex";
        public const string CompNewsIDs = "CompNewsIDs";
        public const string NewsCatName = "NewsCatName";
        public const string NewsCategoryId = "NewsCategoryId";

        public const string StatusID = "StatusID";
        public const string PublicationType = "PublicationType";
        public const string Language = "Language";
        public const string ArticleType = "ArticleType";
        public const string OutPutFile = "OutPutFile";
        
        public const string screenshot_fileName = "screenshot_fileName";
        public const string DA = "DA";
        public const string CoverageCategory = "CoverageCategory";

        public const string AVE = "AVE";
        public const string CirculationScore = "CirculationScore";
        public const string pi_score_Print = "pi_score";
        public const string pi_score_online = "pi_score";


        public const string USP_CoverageDossier_OverView_Page_Data_DT3 = "USP_CoverageDossier_OverView_Page_Data_DT3";
        public const string USP_Fetch_DataByCDID_DT3 = "USP_Fetch_DataByCDID_DT3";
        public const string USP_FetchPending_DCIDsToProcess_DT3 = "USP_FetchPending_DCIDsToProcess_DT3";

        public const string USP_Fetch_Tweets_Data_DT3 = "USP_Fetch_Tweets_Data_DT3";
        public const string USP_Fetch_Print_Data_DT3 = "USP_Fetch_DataByCDID_Print_DT3";
        public const string USP_Fetch_Online_Data_DT3 = "USP_Fetch_DataByCDID_Online_DT3";


        public const string Totalnoof_Tweets = "Totalnoof_Tweets";
        public const string Total_Reach="Total_Reach";
        public const string Total_Engagement="Total_Engagement";
        public const string Totalnoof_Retweets="Totalnoof_Retweets";
        public const string Totalnoof_Participants_Profiles="Totalnoof_Participants_Profiles";
        public const string Totalnoof_Media_and_Influencers="Totalnoof_Media_and_Influencers";
        public const string Totalnoof_threads="Totalnoof_threads";
        public const string Total_Views="Total_Views";

        public const string Totalnoofmediaandinfluencers = "Totalnoofmediaandinfluencers";
        public const string ActiveParticipants = "ActiveParticipants";
        public const string Newparticipants = "Newparticipants";
        public const string Totalthreads = "Totalthreads";

        public const string DGLog_Logid = "Logid";
        public const string DGLog_CDID = "CDID";
        public const string DGLog_FunctionName = "FunctionName";
        public const string DGLog_ERROR = "ERROR";
        public const string DGLog_CreatedAt = "CreatedAt";

        #endregion

    }
}
