using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sahadeva_NewsLetter
{
    public class DataBaseConstants
    {
        #region Connection String

        public const string ConnectionString = "ConnectionString";

        #endregion Connection String


        #region Procedures

        public const string USP_Fetch_NLIDs = "USP_Fetch_NLIDs";
        public const string USP_NL_FetchDataByNLID = "USP_NL_FetchDataByNLID";
        public const string USP_NL_FetchDataToEditByNLID = "USP_NL_FetchDataToEditByNLID";
        public const string USP_NL_UpdateArticleDetails = "USP_NL_UpdateArticleDetails";
        public const string USP_NL_FetchArticleByID = "USP_NL_FetchArticleByID";
        public const string USP_NL_FetchDataByNLID_Dynamic = "USP_NL_FetchDataByNLID_Dynamic";       
     
        

        

        //public const string USP_NL_FetchDataByNLID = "USP_NL_FetchDataByNLID_Mayuresh";
        
        #endregion Procedures

        #region Feilds

        public const string NLID = "NLID";

        public const string ArticleID = "ArticleID";
        public const string MediaTypeID = "MediaTypeID";
        public const string Cluster = "Cluster";
        public const string Headline = "Headline";
        public const string Summary = "Summary";
        public const string Sentiment = "Sentiment";
        public const string Publication = "Publication";
        public const string ArticleType = "ArticleType";


        #endregion
    }
}
