using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using Sahadeva_NewsLetter;

namespace Sahadeva_NL_Edit
{
    public partial class EditArticleDetails : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

                if (Request.QueryString["ArticleID"] != null && Request.QueryString["MediaTypeID"] != null)
                {

                    Int32 ArticleID = Convert.ToInt32(Request.QueryString["ArticleID"]);
                    Int32 MediaTypeID = Convert.ToInt32(Request.QueryString["MediaTypeID"]);

                    BindArticleData(ArticleID, MediaTypeID);
                }
            }
        }

        public void btnUpdate_Click(object sender, EventArgs e)
        {

            try
            {
                Int32 ArticleID = Convert.ToInt32(Request.QueryString["ArticleID"]);
                Int32 MediaTypeID = Convert.ToInt32(Request.QueryString["MediaTypeID"]);


                eArticle objArticle = new eArticle();
                objArticle.ArticleID = ArticleID;
                objArticle.MediaTypeID = MediaTypeID;
                objArticle.Cluster = Convert.ToString(txtCluster.Text);
                objArticle.Headline = Convert.ToString(txtHeadline.Text);
                objArticle.Summary = Convert.ToString(txtSummary.Text);
                objArticle.Sentiment = Convert.ToString(ddlSentiment.SelectedValue);
                objArticle.SentimentColor = Convert.ToString(lblColour.Text);                
                objArticle.Publication = Convert.ToString(txtPublication.Text);
                objArticle.ArticleType = Convert.ToString(txtArticleType.Text);

                new DAL().UpdateArticleDetails(objArticle);

                ScriptManager.RegisterStartupScript(this, this.GetType(), "UpdateSuccessScript", "alert('Update successful!');", true);

                //Response.Redirect("http://localhost:49312/Edit_NL_HTML_Altr_Popup.aspx");

                string script = "window.close();";
                ClientScript.RegisterStartupScript(this.GetType(), "CloseScript", script, true);

            }
            catch (Exception ex)
            {
                
                throw ex;
            }

        }

        public void BindArticleData(Int32 ArticleID, Int32 MediaTypeID)
        {
            try
            {
                DataTable dt = new DAL().FetchArticleByID(ArticleID, MediaTypeID);

                if (dt != null)
                {
                    DataRow dr = dt.Rows[0];
                    if (dt.Rows.Count > 0)
                    {
                        txtCluster.Text = Convert.ToString(dr["Cluster"]);
                        txtHeadline.Text = Convert.ToString(dr["Headline"]);
                        if (ddlSentiment.Items.FindByText(Convert.ToString(dr["Sentiment"])) != null)
                        {
                            ddlSentiment.Items.FindByText(Convert.ToString(dr["Sentiment"])).Selected = true;
                        }
                        lblColour.Text = Convert.ToString(dr["SentimentColor"]);
                        txtSummary.Text = Convert.ToString(dr["Summary"]);
                        txtPublication.Text = Convert.ToString(dr["Publication"]);
                        txtArticleType.Text = Convert.ToString(dr["ArticleType"]);
                    }

                }
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
        }

        protected void ddlSentiment_OnSelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedValue = ddlSentiment.SelectedValue;
            if (selectedValue=="Negative")
            {
                lblColour.Text = "#FF5733";
            }
            else if (selectedValue == "Positive")
            {
                lblColour.Text = "#43b049";
            }
            else if (selectedValue == "Neutral")
            {
                lblColour.Text = "#73c6fb";
            }  
        }
    }
}