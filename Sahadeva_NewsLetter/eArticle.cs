using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Runtime.Serialization;

namespace Sahadeva_NL_Edit
{
    [DataContract]
    public class eArticle
    {

        [DataMember]
        public Int32 ArticleID { get; set; }

        [DataMember]
        public Int32 MediaTypeID { get; set; }

        [DataMember]
        public string Cluster { get; set; }

        [DataMember]
        public string Headline { get; set; }

        [DataMember]
        public string Summary { get; set; }

        [DataMember]
        public string Sentiment { get; set; }

        [DataMember]
        public string SentimentColor { get; set; }

        [DataMember]
        public string Publication { get; set; }

        [DataMember]
        public string ArticleType { get; set; }
    }
}