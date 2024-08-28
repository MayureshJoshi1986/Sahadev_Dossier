using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sahadeva.Dossier.Entities
{    
    public class NewsArticle
    {
        public string HeadLine { get; set; }
        public string NewsText { get; set; }
        public DateTime NewsDate { get; set; }
        public string Edition { get; set; }
        public string Pageno { get; set; }
        public string NewsURL { get; set; }
        public string IMG_URL { get; set; }
        public string Topics { get; set; }
        public string Sentiment { get; set; }
        public string CoverageCategory { get; set; }

        public int Count { get; set; }
        public int Rowno { get; set; }
        public int GroupID { get; set; }
        public string GroupName { get; set; }
        public int CompanyID { get; set; }
        public string CompanyLogoURL { get; set; }
        public string GroupLogoURL { get; set; }
        public int PublicationID { get; set; }
        public string Publication { get; set; }
        public string PublicationLogoURL { get; set; } 
        public string PublicationType { get; set; }
        public string ArticleType { get; set; }         
        
        
        public string News { get; set; }
        public string Physicalpath { get; set; }
        public int CompNewsId { get; set; }
        public string Journalist { get; set; }
        public int nocol { get; set; }
        public int ht_cm { get; set; }
        public string Language { get; set; }
        public int News_id { get; set; }
        public string SubGroup { get; set; }
        public Int32 SubGroupID { get; set; }
        public string NewsPaper { get; set; }
        public string Summary { get; set; }
        public string NewsGroup { get; set; }
        public string NewsSubGroup { get; set; }
        public int NewsGroupID { get; set; }

        public string AVE { get; set; }
        public string CirculationScore { get; set; }
        public string pi_score { get; set; }
    }
}
