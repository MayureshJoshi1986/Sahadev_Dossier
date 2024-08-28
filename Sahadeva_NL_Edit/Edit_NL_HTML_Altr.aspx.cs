using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Sahadeva_NewsLetter;
using System.Data;

namespace Sahadeva_NL_Edit
{
    public partial class Edit_NL_HTML_Altr : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                DataSet ds = new DAL().FetchDataToEditByNLID(2);

                DataTable dt1 = ds.Tables[0];
                lblEntityName.Text = Convert.ToString(dt1.Rows[0]["EntityName"]);
                lblNewsLetterTitle.Text = Convert.ToString(dt1.Rows[0]["Title"]);
                lblNewsLetterDate.Text = Convert.ToString(dt1.Rows[0]["Date"]);

                BindArticleData(ds);
            }
        }

        public void BindArticleData(DataSet ds)
        {
            DataTable dt1 = ds.Tables[1];

            int Count = Convert.ToInt32(dt1.Rows[0]["Count_Sections"]);

            if (Count > 0)
            {
                DataTable dt2 = ds.Tables[2];
                lblEntityNameSection_1.Text = Convert.ToString(dt2.Rows[0]["SectionType"]) + " - " + Convert.ToString(dt2.Rows[0]["EntityName"]);
                lblEntityDescription_1.Text = Convert.ToString(dt2.Rows[0]["Summary"]);

                DataTable dt3 = ds.Tables[3];
                if (dt3 != null)
                {
                    if (dt3.Rows.Count > 0)
                    {
                        lblSectionTitleOnline_1.Text = Convert.ToString(dt3.Rows[0]["SectionTitle"]);
                        lvOnlineArticles_1.DataSource = dt3;
                        lvOnlineArticles_1.DataBind();
                    }

                }

                DataTable dt4 = ds.Tables[4];

                if (dt4 != null)
                {
                    if (dt4.Rows.Count > 0)
                    {
                        lblSectionTitlePrint_1.Text = Convert.ToString(dt4.Rows[0]["SectionTitle"]);
                        lvPrintArticles_1.DataSource = dt4;
                        lvPrintArticles_1.DataBind();
                    }
                }
            }
        }

        protected void lvOnlineArticles_1_ItemEditing(object sender, ListViewEditEventArgs e)
        {
            lvOnlineArticles_1.EditIndex = e.NewEditIndex;

            DataSet ds = new DAL().FetchDataToEditByNLID(2);

            BindArticleData(ds);
        }

        protected void lvOnlineArticles_1_ItemDeleting(object sender, ListViewDeleteEventArgs e)
        {
            lvOnlineArticles_1.EditIndex = -1;
            DataSet ds = new DAL().FetchDataToEditByNLID(2);
            BindArticleData(ds);
            e.Cancel = true; // Cancel the default deletion behavior
        }

        protected void lvOnlineArticles_1_ItemDataBound(object sender, ListViewItemEventArgs e)
        {
            if (e.Item.ItemType == ListViewItemType.DataItem)
            {
                Button btnDelete = (Button)e.Item.FindControl("btnDelete_Online");
                if (btnDelete != null)
                {
                    btnDelete.OnClientClick = "return confirm('Are you sure you want to delete this record?');";
                }
            }
        }

        protected void lvOnlineArticles_1_ItemUpdating(object sender, ListViewUpdateEventArgs e)
        {

            int ArticleID = Convert.ToInt32(lvOnlineArticles_1.DataKeys[e.ItemIndex]["ArticleID"]);
            int MediaTypeID = Convert.ToInt32(lvOnlineArticles_1.DataKeys[e.ItemIndex]["MediaTypeID"]);

            ListViewItem item = lvOnlineArticles_1.EditItem;

            TextBox txtCluster_Online = (TextBox)item.FindControl("txtCluster_Online");
            TextBox txtHeadline_Online = (TextBox)item.FindControl("txtHeadline_Online");
            TextBox txtSummary_Online = (TextBox)item.FindControl("txtSummary_Online");
            TextBox txtSentiment_Online = (TextBox)item.FindControl("txtSentiment_Online");
            TextBox txtPublication_Online = (TextBox)item.FindControl("txtPublication_Online");
            TextBox txtArticleType_Online = (TextBox)item.FindControl("txtArticleType_Online");

            eArticle objArticle = new eArticle();
            objArticle.ArticleID = ArticleID;
            objArticle.MediaTypeID = MediaTypeID;
            objArticle.Cluster = Convert.ToString(txtCluster_Online.Text);
            objArticle.Headline = Convert.ToString(txtHeadline_Online.Text);
            objArticle.Summary = Convert.ToString(txtSummary_Online.Text);
            objArticle.Sentiment = Convert.ToString(txtSentiment_Online.Text);
            objArticle.Publication = Convert.ToString(txtPublication_Online.Text);
            objArticle.ArticleType = Convert.ToString(txtArticleType_Online.Text);

            new DAL().UpdateArticleDetails(objArticle);

            new DAL().FetchDataToEditByNLID(2);

            lvOnlineArticles_1.EditIndex = -1;
            DataSet ds = new DAL().FetchDataToEditByNLID(2);
            BindArticleData(ds);
            e.Cancel = true; // Cancel the default updating behavior
        }

        protected void lvOnlineArticles_1_ItemCanceling(object sender, ListViewCancelEventArgs e)
        {
            lvOnlineArticles_1.EditIndex = -1;

            DataSet ds = new DAL().FetchDataToEditByNLID(2);
            BindArticleData(ds);
        }

    }
}