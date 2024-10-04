namespace Sahadeva.Dossier.Common
{
    public class DatabaseConstants
    {
        public const string ConnectionString = "ConnectionString";
        
        // SPs
        public const string USP_FetchPending_DCIDsToProcess_All = "FetchPending_DCIDsToProcess_All_NEW";
        public const string USP_CoverageDossier_UpdateStatus = "USP_CoverageDossier_UpdateStatus";

        // SP Params
        public const string CDID = "CDID";
        public const string StatusID = "StatusID";
    }
}
