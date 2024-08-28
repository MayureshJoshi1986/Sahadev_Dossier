using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.Common;
using Sahadeva_NL_Edit;

namespace Sahadeva_NewsLetter
{
    public class DAL
    {
        public DataTable Fetch_NLIDs()
        {
            try
            {
                DataTable dt = new DataTable();

                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DataBaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DataBaseConstants.USP_Fetch_NLIDs))
                    {
                        #region DataSet
                        using (DataSet ds = DataAccessWrapper.ExecuteDataSet(dbCommand))
                        {
                            if (ds != null)
                            {
                                dt = ds.Tables[0];
                            }
                        }
                        #endregion
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        public DataSet FetchDataByNLID(Int32 NLID)
        {
            try
            {
                DataSet ds = new DataSet();

                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DataBaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DataBaseConstants.USP_NL_FetchDataByNLID))
                    //using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DataBaseConstants.USP_NL_FetchDataByNLID_Dynamic))
                    {
                        #region DataSet

                        DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.NLID, DbType.Int32, NLID);
                        ds = DataAccessWrapper.ExecuteDataSet(dbCommand);

                        #endregion
                    }
                }
                return ds;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataSet FetchDataToEditByNLID(Int32 NLID)
        {
            try
            {
                DataSet ds = new DataSet();

                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DataBaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DataBaseConstants.USP_NL_FetchDataToEditByNLID))
                    {
                        #region DataSet

                        DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.NLID, DbType.Int32, NLID);
                        ds = DataAccessWrapper.ExecuteDataSet(dbCommand);

                        #endregion
                    }
                }
                return ds;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public void UpdateArticleDetails(eArticle objArticle )
        {
            try
            {
                DataSet ds = new DataSet();

                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DataBaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DataBaseConstants.USP_NL_UpdateArticleDetails))
                    {
                        #region DataSet

                        DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.ArticleID, DbType.Int32, objArticle.ArticleID);
                        DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.MediaTypeID, DbType.Int32, objArticle.MediaTypeID);

                        if (objArticle.Cluster != null)
                        { DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.Cluster, DbType.String, objArticle.Cluster); }

                        if (objArticle.Headline != null)
                        { DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.Headline, DbType.String, objArticle.Headline); }

                        if (objArticle.Summary != null)
                        { DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.Summary, DbType.String, objArticle.Summary); }

                        if (objArticle.Sentiment != null)
                        { DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.Sentiment, DbType.String, objArticle.Sentiment); }

                        if (objArticle.Publication != null)
                        { DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.Publication, DbType.String, objArticle.Publication); }

                        if (objArticle.ArticleType != null)
                        { DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.ArticleType, DbType.String, objArticle.ArticleType); }


                        DataAccessWrapper.ExecuteDataSet(dbCommand);

                        #endregion
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable FetchArticleByID(Int32 ArticleID,Int32 MediaTypeID)
        {
            try
            {
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();


                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DataBaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DataBaseConstants.USP_NL_FetchArticleByID))
                    {
                        #region DataSet

                        DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.ArticleID, DbType.Int32, ArticleID);
                        DataAccessWrapper.AddInParameter(dbCommand, DataBaseConstants.MediaTypeID, DbType.Int32, MediaTypeID);

                        ds = DataAccessWrapper.ExecuteDataSet(dbCommand);
                        dt = ds.Tables[0];
                        #endregion
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
