using Sahadeva.Dossier.Common;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;

namespace Sahadeva.Dossier.DAL
{
    public enum DossierDataSet
    {
        ClientData,
        OverviewTable,
        OverviewSummary,
        PrintTable,
        OnlineTable,
        PrintDetail,
        OnlineDetail,
    }

    public class DossierDAL
    {
        private readonly Dictionary<DossierDataSet, string> _dataSetMap = new Dictionary<DossierDataSet, string> 
        {
            { DossierDataSet.ClientData, "Fetch_CoverPage_Section" },
            { DossierDataSet.OverviewTable, "Fetch_OverviewTable_Section" },
            { DossierDataSet.OverviewSummary, "Fetch_OverviewSummary_Section" },
            { DossierDataSet.PrintTable, "Fetch_TableContentPrint_Section" },
            { DossierDataSet.OnlineTable, "Fetch_TableContentOnline_Section" },
            { DossierDataSet.PrintDetail, "Fetch_ArticleDetailPrint_Section" },
            { DossierDataSet.OnlineDetail, "Fetch_ArticleDetailOnline_Section" },
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
    }
}
