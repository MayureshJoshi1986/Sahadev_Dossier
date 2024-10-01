using Sahadeva.Dossier.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;

namespace Sahadeva.Dossier.DAL
{
    public enum DossierDataSet
    {
        CoverPage = 1,
        OverviewTable,
        OverviewSummary,
        ClientGraph,
        TableContentPrint,
        ArticleDetailPrint,
        TableContentOnline,
        ArticleDetailOnline,
        CompetitorGraph
    }

    public class DossierDAL
    {
        private readonly Dictionary<DossierDataSet, string> _dataSetMap = new Dictionary<DossierDataSet, string> 
        {
            { DossierDataSet.CoverPage, "Fetch_CoverPage_Section" },
            { DossierDataSet.OverviewTable, "Fetch_OverviewTable_Section" },
            { DossierDataSet.OverviewSummary, "Fetch_OverviewSummary_Section" },
            { DossierDataSet.ClientGraph, "Fetch_ClientGraph_Section" },
            { DossierDataSet.TableContentPrint, "Fetch_TableContentPrint_Section" },
            { DossierDataSet.TableContentOnline, "Fetch_TableContentOnline_Section" },
            { DossierDataSet.ArticleDetailPrint, "Fetch_ArticleDetailPrint_Section" },
            { DossierDataSet.ArticleDetailOnline, "Fetch_ArticleDetailOnline_Section" },
            { DossierDataSet.CompetitorGraph, "Fetch_CompetitorGraph_Section" }
        };

        public DataTable FetchData(int coverageDossierId, DossierDataSet dataSet)
        {
            DataSet ds;

            using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
            {
                using (DbCommand dbcommand = DataAccessWrapper.GetStoredProcCommand(_dataSetMap[dataSet]))
                {
                    DataAccessWrapper.AddInParameter(dbcommand, DatabaseConstants.CDID, DbType.Int32, coverageDossierId);

                    ds = DataAccessWrapper.ExecuteDataSet(dbcommand);
                }
            }

            return ds.Tables[0];
        }

        public DataTable FetchPending_DCIDsToProcess_All()
        {
            DataTable dt = new DataTable();

            try
            {
                using (DataAccessWrapper DataAccessWrapper = new DataAccessWrapper(DatabaseConstants.ConnectionString))
                {
                    using (DbCommand dbCommand = DataAccessWrapper.GetStoredProcCommand(DatabaseConstants.USP_FetchPending_DCIDsToProcess_All))
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
    }
}
