using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Sahadeva.Dossier.Entities;
using Sahadeva.Dossier.Common;
using System.Data.Common;
using System.Data;

namespace Sahadeva.Dossier.DAL
{
    public class DossierTemp3DAL
    {
        public DataTable FetchPending_DCIDsToProcess_DT3()
        {
            DataTable dt = new DataTable();

            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    //using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Dossier_FetchPrintArticleData))
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_FetchPending_DCIDsToProcess_DT3))
                    {
                        #region DataSet
                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbCommand))
                        {
                            if (dsiObject != null)
                            {
                                if (dsiObject.Tables[0] != null)
                                {
                                    dt = dsiObject.Tables[0];
                                }

                            }
                        }
                        #endregion
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

        public DataTable Dossier_FetchClientData_DT3(Int32 CDID)
        {
            DataTable dt = new DataTable();

            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    //using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Dossier_FetchPrintArticleData))
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Fetch_DataByCDID_DT3))
                    {
                        dbCommand.CommandTimeout = 0;
                        DataAccessWrapper.AddInParameter(dbCommand, DatabaseConstants.CDID, DbType.Int32, CDID);

                        #region DataSet
                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbCommand))
                        {
                            if (dsiObject != null)
                            {
                                if (dsiObject.Tables[0] != null)
                                {
                                    dt = dsiObject.Tables[0];
                                }

                            }
                        }
                        #endregion
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }

        public List<NewsArticle> Dossier_FetchArticleData_Print_DT3(Int32 CDID)
        {
            List<NewsArticle> lstNewsArticle = new List<NewsArticle>();
            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    //using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Dossier_FetchPrintArticleData))
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Fetch_DataByCDID_DT3))
                    {
                        dbCommand.CommandTimeout = 0;
                        DataAccessWrapper.AddInParameter(dbCommand, DatabaseConstants.CDID, DbType.Int32, CDID);

                        #region DataSet
                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbCommand))
                        {
                            if (dsiObject != null)
                            {
                                DataTable dtiObject = dsiObject.Tables[1];
                                NewsArticle NewsArticle = null;
                                foreach (DataRow driObject in dtiObject.Rows)
                                {
                                    NewsArticle = new NewsArticle();

                                    if (!driObject[DatabaseConstants.headline].Equals(DBNull.Value))
                                    {
                                        NewsArticle.HeadLine = Convert.ToString(driObject[DatabaseConstants.headline]);
                                    }
                                    if (!driObject[DatabaseConstants.NewsText].Equals(DBNull.Value))
                                    {
                                        NewsArticle.NewsText = Convert.ToString(driObject[DatabaseConstants.NewsText]);
                                    }
                                    if (!driObject[DatabaseConstants.News_dt].Equals(DBNull.Value))
                                    {
                                        NewsArticle.NewsDate = Convert.ToDateTime(driObject[DatabaseConstants.News_dt]);
                                    }
                                    if (!driObject[DatabaseConstants.edition].Equals(DBNull.Value))
                                    {
                                        NewsArticle.Edition = Convert.ToString(driObject[DatabaseConstants.edition]);
                                    }
                                    if (!driObject[DatabaseConstants.page].Equals(DBNull.Value))
                                    {
                                        NewsArticle.Pageno = Convert.ToString(driObject[DatabaseConstants.page]);
                                    }
                                    if (!driObject[DatabaseConstants.img_url].Equals(DBNull.Value))
                                    {
                                        NewsArticle.NewsURL = Convert.ToString(driObject[DatabaseConstants.img_url]);
                                    }
                                    if (!driObject[DatabaseConstants.News_dt].Equals(DBNull.Value))
                                    {
                                        NewsArticle.NewsDate = Convert.ToDateTime(driObject[DatabaseConstants.News_dt]);
                                    }
                                    if (!driObject[DatabaseConstants.NewsURL].Equals(DBNull.Value))
                                    {
                                        NewsArticle.NewsURL = Convert.ToString(driObject[DatabaseConstants.NewsURL]);
                                    }
                                    if (!driObject[DatabaseConstants.img_url].Equals(DBNull.Value))
                                    {
                                        NewsArticle.IMG_URL = Convert.ToString(driObject[DatabaseConstants.img_url]);
                                    }
                                    if (!driObject[DatabaseConstants.Topics].Equals(DBNull.Value))
                                    {
                                        NewsArticle.Topics = Convert.ToString(driObject[DatabaseConstants.Topics]);
                                    }
                                    if (!driObject[DatabaseConstants.Sentiment].Equals(DBNull.Value))
                                    {
                                        NewsArticle.Sentiment = Convert.ToString(driObject[DatabaseConstants.Sentiment]);
                                    }
                                    if (!driObject[DatabaseConstants.Publication].Equals(DBNull.Value))
                                    {
                                        NewsArticle.Publication = Convert.ToString(driObject[DatabaseConstants.Publication]);
                                    }
                                    if (!driObject[DatabaseConstants.PublicationType].Equals(DBNull.Value))
                                    {
                                        NewsArticle.PublicationType = Convert.ToString(driObject[DatabaseConstants.PublicationType]);
                                    }
                                    if (!driObject[DatabaseConstants.Journalist].Equals(DBNull.Value))
                                    {
                                        NewsArticle.Journalist = Convert.ToString(driObject[DatabaseConstants.Journalist]);
                                    }
                                    if (!driObject[DatabaseConstants.Language].Equals(DBNull.Value))
                                    {
                                        NewsArticle.Language = Convert.ToString(driObject[DatabaseConstants.Language]);
                                    }
                                    if (!driObject[DatabaseConstants.ArticleType].Equals(DBNull.Value))
                                    {
                                        NewsArticle.ArticleType = Convert.ToString(driObject[DatabaseConstants.ArticleType]);
                                    }

                                    if (!driObject[DatabaseConstants.CoverageCategory].Equals(DBNull.Value))
                                    {
                                        NewsArticle.CoverageCategory = Convert.ToString(driObject[DatabaseConstants.CoverageCategory]);
                                    }

                                    if (!driObject[DatabaseConstants.AVE].Equals(DBNull.Value))
                                    {
                                        NewsArticle.AVE = Convert.ToString(driObject[DatabaseConstants.AVE]);
                                    }
                                    if (!driObject[DatabaseConstants.CirculationScore].Equals(DBNull.Value))
                                    {
                                        NewsArticle.CirculationScore = Convert.ToString(driObject[DatabaseConstants.CirculationScore]);
                                    }
                                    if (!driObject[DatabaseConstants.pi_score_Print].Equals(DBNull.Value))
                                    {
                                        NewsArticle.pi_score = Convert.ToString(driObject[DatabaseConstants.pi_score_Print]);
                                    }

                                    lstNewsArticle.Add(NewsArticle);
                                }
                            }
                        }
                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return lstNewsArticle;
        }

        public List<NewsArticle_Online> Dossier_FetchArticleData_Online_DT3(Int32 CDID)
        {
            List<NewsArticle_Online> lstNewsArticle = new List<NewsArticle_Online>();
            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Fetch_DataByCDID_DT3))
                    {

                        DataAccessWrapper.AddInParameter(dbCommand, DatabaseConstants.CDID, DbType.Int32, CDID);
                        dbCommand.CommandTimeout = 0;
                        #region DataSet
                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbCommand))
                        {
                            if (dsiObject != null)
                            {
                                if (dsiObject.Tables.Count > 0)
                                {
                                    DataTable dtiObject = dsiObject.Tables[2];
                                    NewsArticle_Online NewsArticle = null;
                                    foreach (DataRow driObject in dtiObject.Rows)
                                    {
                                        NewsArticle = new NewsArticle_Online();

                                        if (!driObject[DatabaseConstants.Title].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Title = Convert.ToString(driObject[DatabaseConstants.Title]);
                                        }
                                        if (!driObject[DatabaseConstants.NewsText].Equals(DBNull.Value))
                                        {
                                            NewsArticle.NewsText = Convert.ToString(driObject[DatabaseConstants.NewsText]);
                                        }
                                        if (!driObject[DatabaseConstants.News_dt].Equals(DBNull.Value))
                                        {
                                            NewsArticle.NewsDate = Convert.ToDateTime(driObject[DatabaseConstants.News_dt]);
                                        }
                                        if (!driObject[DatabaseConstants.edition].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Edition = Convert.ToString(driObject[DatabaseConstants.edition]);
                                        }
                                        if (!driObject[DatabaseConstants.img_url].Equals(DBNull.Value))
                                        {
                                            NewsArticle.IMG_URL = Convert.ToString(driObject[DatabaseConstants.img_url]);
                                        }
                                        if (!driObject[DatabaseConstants.News_dt].Equals(DBNull.Value))
                                        {
                                            NewsArticle.NewsDate = Convert.ToDateTime(driObject[DatabaseConstants.News_dt]);
                                        }
                                        if (!driObject[DatabaseConstants.NewsURL].Equals(DBNull.Value))
                                        {
                                            NewsArticle.NewsURL = Convert.ToString(driObject[DatabaseConstants.NewsURL]);
                                        }

                                        if (!driObject[DatabaseConstants.Topics].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Topics = Convert.ToString(driObject[DatabaseConstants.Topics]);
                                        }
                                        if (!driObject[DatabaseConstants.Sentiment].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Sentiment = Convert.ToString(driObject[DatabaseConstants.Sentiment]);
                                        }
                                        if (!driObject[DatabaseConstants.Publication].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Publication = Convert.ToString(driObject[DatabaseConstants.Publication]);
                                        }

                                        if (!driObject[DatabaseConstants.Cluster].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Cluster = Convert.ToString(driObject[DatabaseConstants.Cluster]);
                                        }
                                        if (!driObject[DatabaseConstants.NewsTraffic].Equals(DBNull.Value))
                                        {
                                            NewsArticle.NewsTraffic = Convert.ToString(driObject[DatabaseConstants.NewsTraffic]);
                                        }
                                        if (!driObject[DatabaseConstants.PublicationType].Equals(DBNull.Value))
                                        {
                                            NewsArticle.PublicationType = Convert.ToString(driObject[DatabaseConstants.PublicationType]);
                                        }
                                        if (!driObject[DatabaseConstants.Journalist].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Journalist = Convert.ToString(driObject[DatabaseConstants.Journalist]);
                                        }
                                        if (!driObject[DatabaseConstants.Language].Equals(DBNull.Value))
                                        {
                                            NewsArticle.Language = Convert.ToString(driObject[DatabaseConstants.Language]);
                                        }
                                        if (!driObject[DatabaseConstants.ArticleType].Equals(DBNull.Value))
                                        {
                                            NewsArticle.ArticleType = Convert.ToString(driObject[DatabaseConstants.ArticleType]);
                                        }
                                        if (!driObject[DatabaseConstants.screenshot_fileName].Equals(DBNull.Value))
                                        {
                                            NewsArticle.screenshot_fileName = Convert.ToString(driObject[DatabaseConstants.screenshot_fileName]);
                                        }
                                        if (!driObject[DatabaseConstants.DA].Equals(DBNull.Value))
                                        {
                                            NewsArticle.DA = Convert.ToString(driObject[DatabaseConstants.DA]);
                                        }

                                        if (!driObject[DatabaseConstants.CoverageCategory].Equals(DBNull.Value))
                                        {
                                            NewsArticle.CoverageCategory = Convert.ToString(driObject[DatabaseConstants.CoverageCategory]);
                                        }


                                        if (!driObject[DatabaseConstants.pi_score_online].Equals(DBNull.Value))
                                        {
                                            NewsArticle.pi_score = Convert.ToString(driObject[DatabaseConstants.pi_score_online]);
                                        }
                                        if (!driObject[DatabaseConstants.CoverageCategory].Equals(DBNull.Value))
                                        {
                                            NewsArticle.CoverageCategory = Convert.ToString(driObject[DatabaseConstants.CoverageCategory]);
                                        }
                                        if (!driObject[DatabaseConstants.CoverageCategory].Equals(DBNull.Value))
                                        {
                                            NewsArticle.CoverageCategory = Convert.ToString(driObject[DatabaseConstants.CoverageCategory]);
                                        }

                                        //if (!driObject[DatabaseConstants.edition].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.Edition = Convert.ToString(driObject[DatabaseConstants.edition]);
                                        //}
                                        //if (!driObject[DatabaseConstants.newspaper].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.Publication = Convert.ToString(driObject[DatabaseConstants.newspaper]);
                                        //}
                                        //if (!driObject[DatabaseConstants.NewsPaperid].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.PublicationID = Convert.ToInt32(driObject[DatabaseConstants.NewsPaperid]);
                                        //}
                                        //if (!driObject[DatabaseConstants.headline].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.HeadLine = Convert.ToString(driObject[DatabaseConstants.headline]);
                                        //}
                                        //if (!driObject[DatabaseConstants.Journalist].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.Journalist = Convert.ToString(driObject[DatabaseConstants.Journalist]);
                                        //}
                                        //if (!driObject[DatabaseConstants.summary].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.News = Convert.ToString(driObject[DatabaseConstants.summary]);
                                        //}
                                        //if (!driObject[DatabaseConstants.img_url].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.NewsURL = Convert.ToString(driObject[DatabaseConstants.img_url]);
                                        //}
                                        //if (!driObject[DatabaseConstants.no_col].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.nocol = Convert.ToInt32(driObject[DatabaseConstants.no_col]);
                                        //}
                                        //if (!driObject[DatabaseConstants.ht_cm].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.ht_cm = Convert.ToInt32(driObject[DatabaseConstants.ht_cm]);
                                        //}
                                        //if (!driObject[DatabaseConstants.LanguageName].Equals(DBNull.Value))
                                        //{
                                        //    NewsArticle.Language = Convert.ToString(driObject[DatabaseConstants.LanguageName]);
                                        //}

                                        lstNewsArticle.Add(NewsArticle);
                                    }
                                }

                            }
                        }
                        #endregion
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            return lstNewsArticle;
        }
        public List<Twitter> Dossier_FetchArticleData_TwitterData_DT3(Int32 CDID)
        {
            List<Twitter> lstTwitter = new List<Twitter>();
            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Fetch_DataByCDID_DT3))
                    {

                        DataAccessWrapper.AddInParameter(dbCommand, DatabaseConstants.CDID, DbType.Int32, CDID);
                        dbCommand.CommandTimeout = 0;
                        #region DataSet
                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbCommand))
                        {
                            if (dsiObject != null)
                            {
                                if (dsiObject.Tables.Count > 0)
                                {
                                    DataTable dtiObject = dsiObject.Tables[1];
                                    Twitter TwitterData = null;
                                    foreach (DataRow driObject in dtiObject.Rows)
                                    {
                                        TwitterData = new Twitter();

                                        if (!driObject[DatabaseConstants.Totalnoof_Tweets].Equals(DBNull.Value))
                                        {
                                            TwitterData.Totalnoof_Tweets = Convert.ToString(driObject[DatabaseConstants.Totalnoof_Tweets]);
                                        }
                                        if (!driObject[DatabaseConstants.Total_Reach].Equals(DBNull.Value))
                                        {
                                            TwitterData.Total_Reach = Convert.ToString(driObject[DatabaseConstants.Total_Reach]);
                                        }
                                        if (!driObject[DatabaseConstants.Total_Engagement].Equals(DBNull.Value))
                                        {
                                            TwitterData.Total_Engagement = Convert.ToString(driObject[DatabaseConstants.Total_Engagement]);
                                        }
                                        if (!driObject[DatabaseConstants.Totalnoof_Retweets].Equals(DBNull.Value))
                                        {
                                            TwitterData.Totalnoof_Retweets = Convert.ToString(driObject[DatabaseConstants.Totalnoof_Retweets]);
                                        }
                                        if (!driObject[DatabaseConstants.Totalnoof_Participants_Profiles].Equals(DBNull.Value))
                                        {
                                            TwitterData.Totalnoof_Participants_Profiles = Convert.ToString(driObject[DatabaseConstants.Totalnoof_Participants_Profiles]);
                                        }
                                        if (!driObject[DatabaseConstants.Totalnoof_Media_and_Influencers].Equals(DBNull.Value))
                                        {
                                            TwitterData.Totalnoof_Media_and_Influencers = Convert.ToString(driObject[DatabaseConstants.Totalnoof_Media_and_Influencers]);
                                        }
                                        if (!driObject[DatabaseConstants.Totalnoof_threads].Equals(DBNull.Value))
                                        {
                                            TwitterData.Totalnoof_threads = Convert.ToString(driObject[DatabaseConstants.Totalnoof_threads]);
                                        }
                                        if (!driObject[DatabaseConstants.Total_Views].Equals(DBNull.Value))
                                        {
                                            TwitterData.Total_Views = Convert.ToString(driObject[DatabaseConstants.Total_Views]);
                                        }
                                        if (!driObject[DatabaseConstants.Totalnoofmediaandinfluencers].Equals(DBNull.Value))
                                        {
                                            TwitterData.Totalnoofmediaandinfluencers = Convert.ToString(driObject[DatabaseConstants.Totalnoofmediaandinfluencers]);
                                        }
                                        if (!driObject[DatabaseConstants.ActiveParticipants].Equals(DBNull.Value))
                                        {
                                            TwitterData.ActiveParticipants = Convert.ToString(driObject[DatabaseConstants.ActiveParticipants]);
                                        }
                                        if (!driObject[DatabaseConstants.Newparticipants].Equals(DBNull.Value))
                                        {
                                            TwitterData.Newparticipants = Convert.ToString(driObject[DatabaseConstants.Newparticipants]);
                                        }
                                        if (!driObject[DatabaseConstants.Totalthreads].Equals(DBNull.Value))
                                        {
                                            TwitterData.Totalthreads = Convert.ToString(driObject[DatabaseConstants.Totalthreads]);
                                        }



                                        lstTwitter.Add(TwitterData);
                                    }
                                }

                            }
                        }
                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return lstTwitter;
        }

        public DataSet CoverageDossier_OverView_Page_DT3(Int32 CDID)
        {
            try
            {
                DataSet ds = new DataSet();
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_CoverageDossier_OverView_Page_Data_DT3))
                    {
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.CDID, DbType.Int32, CDID);

                        DataSet dt = DataAccessWrapper.ExecuteDataSet(dbcommand);

                        ds = dt;
                    }
                }
                return ds;

            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CoverageDossier_UpdateStatus(Int32 CDID, Int32 StatusID)
        {
            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_CoverageDossier_UpdateStatus))
                    {
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.CDID, DbType.Int32, CDID);
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.StatusID, DbType.Int32, StatusID);

                        DataAccessWrapper.ExecuteNonQuery(dbcommand);
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Update_CD_OutputFileLink(Int32 CDID, string OutPutFileLink)
        {
            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Update_CD_OutputFileLink))
                    {
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.CDID, DbType.Int32, CDID);
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.OutPutFile, DbType.String, OutPutFileLink);

                        DataAccessWrapper.ExecuteNonQuery(dbcommand);
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataSet Fetch_Tweets_Data_DT3(int CDID)
        {
            try
            {
                DataSet result = new DataSet();
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Fetch_Tweets_Data_DT3))
                    {
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.CDID, DbType.Int32, CDID);

                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbcommand))
                        {
                            if (dsiObject != null)
                            {
                                foreach (DataTable table in dsiObject.Tables)
                                {
                                    result.Tables.Add(table.Copy());
                                }
                            }
                        }
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataSet Fetch_Print_Data_DT3(int CDID)
        {
            try
            {
                DataSet result = new DataSet();
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Fetch_Print_Data_DT3))
                    {
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.CDID, DbType.Int32, CDID);

                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbcommand))
                        {
                            if (dsiObject != null)
                            {
                                foreach (DataTable table in dsiObject.Tables)
                                {
                                    result.Tables.Add(table.Copy());
                                }
                            }
                        }
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
         public DataSet Fetch_Online_Data_DT3(int CDID)
        {
            try
            {
                DataSet result = new DataSet();
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Fetch_Online_Data_DT3))
                    {
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.CDID, DbType.Int32, CDID);

                        using (DataSet dsiObject = DataAccessWrapper.ExecuteDataSet(dbcommand))
                        {
                            if (dsiObject != null)
                            {
                                foreach (DataTable table in dsiObject.Tables)
                                {
                                    result.Tables.Add(table.Copy());
                                }
                            }
                        }
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
      
        public void DossierGanrationLog(Int32 CDID, string FunctionName, string ERROR)
        {
            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_Insert_DossierGanrationLog))
                    {
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.DGLog_CDID, DbType.Int32, CDID);
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.DGLog_FunctionName, DbType.String, FunctionName);
                        DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.DGLog_ERROR, DbType.String, ERROR);

                        DataAccessWrapper.ExecuteNonQuery(dbcommand);
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
