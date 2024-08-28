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
    public partial class Edit_NL_HTML : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {
                DataSet ds = new DAL().FetchDataToEditByNLID(2);

                DataTable dt1 = ds.Tables[0];
                lblEntityName.Text =Convert.ToString(dt1.Rows[0]["EntityName"]);
                lblNewsLetterTitle.Text = Convert.ToString(dt1.Rows[0]["Title"]);
                lblNewsLetterDate.Text = Convert.ToString(dt1.Rows[0]["Date"]);

                BindArticleData(ds);
            }            
        }

        public void BindArticleData(DataSet ds)
        {
            DataTable dt1 = ds.Tables[1];

            int Count = Convert.ToInt32(dt1.Rows[0]["Count_Sections"]);

            if (Count>0)
            {
                DataTable dt2 = ds.Tables[2];
                lblEntityNameSection_1.Text = Convert.ToString(dt2.Rows[0]["SectionType"]) + " - " + Convert.ToString(dt2.Rows[0]["EntityName"]);
                lblEntityDescription_1.Text = Convert.ToString(dt2.Rows[0]["Summary"]);

                DataTable dt3 = ds.Tables[3];
                if (dt3!=null)
                {
                    if (dt3.Rows.Count>0)
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

            if (Count > 1)
            {
                DataTable dt5 = ds.Tables[5];
                lblEntityNameSection_2.Text = Convert.ToString(dt5.Rows[0]["SectionType"]) + " - " + Convert.ToString(dt5.Rows[0]["EntityName"]);
                lblEntityDescription_2.Text = Convert.ToString(dt5.Rows[0]["Summary"]);

                DataTable dt6 = ds.Tables[6];

                if (dt6 != null)
                {
                    if (dt6.Rows.Count > 0)
                    {
                        lblSectionTitleOnline_2.Text = Convert.ToString(dt6.Rows[0]["SectionTitle"]);
                        lvOnlineArticles_2.DataSource = dt6;
                        lvOnlineArticles_2.DataBind();
                    }
                }

                DataTable dt7 = ds.Tables[7];

                if (dt7 != null)
                {
                    if (dt7.Rows.Count > 0)
                    {
                        lblSectionTitlePrint_2.Text = Convert.ToString(dt7.Rows[0]["SectionTitle"]);
                        lvPrintArticles_2.DataSource = dt7;
                        lvPrintArticles_2.DataBind();
                    }
                }
            }

            if (Count > 2)
            {
                DataTable dt8 = ds.Tables[8];
                lblEntityNameSection_3.Text = Convert.ToString(dt8.Rows[0]["SectionType"]) + " - " + Convert.ToString(dt8.Rows[0]["EntityName"]);
                lblEntityDescription_3.Text = Convert.ToString(dt8.Rows[0]["Summary"]);

                DataTable dt9 = ds.Tables[9];
                if (dt9 != null)
                {
                    if (dt9.Rows.Count > 0)
                    {
                        lblSectionTitleOnline_3.Text = Convert.ToString(dt9.Rows[0]["SectionTitle"]);
                        lvOnlineArticles_3.DataSource = dt9;
                        lvOnlineArticles_3.DataBind();
                    }
                }
                DataTable dt10 = ds.Tables[10];
                if (dt10 != null)
                {
                    if (dt10.Rows.Count > 0)
                    {
                        lblSectionTitlePrint_3.Text = Convert.ToString(dt10.Rows[0]["SectionTitle"]);
                        lvPrintArticles_3.DataSource = dt10;
                        lvPrintArticles_3.DataBind();
                    }
                }
            }

            if (Count > 3)
            {
                DataTable dt11 = ds.Tables[11];
                lblEntityNameSection_4.Text = Convert.ToString(dt11.Rows[0]["SectionType"]) + " - " + Convert.ToString(dt11.Rows[0]["EntityName"]);
                lblEntityDescription_4.Text = Convert.ToString(dt11.Rows[0]["Summary"]);

                DataTable dt12 = ds.Tables[12];
                if (dt12 != null)
                {
                    if (dt12.Rows.Count > 0)
                    {
                        lblSectionTitleOnline_4.Text = Convert.ToString(dt12.Rows[0]["SectionTitle"]);
                        lvOnlineArticles_4.DataSource = dt12;
                        lvOnlineArticles_4.DataBind();
                    }
                }
                DataTable dt13 = ds.Tables[13];
                if (dt13 != null)
                {
                    if (dt13.Rows.Count > 0)
                    {
                        lblSectionTitlePrint_4.Text = Convert.ToString(dt13.Rows[0]["SectionTitle"]);
                        lvPrintArticles_4.DataSource = dt13;
                        lvPrintArticles_4.DataBind();
                    }
                }
            }

            if (Count > 4)
            {
                DataTable dt14 = ds.Tables[14];
                lblEntityNameSection_5.Text = Convert.ToString(dt14.Rows[0]["SectionType"]) + " - " + Convert.ToString(dt14.Rows[0]["EntityName"]);
                lblEntityDescription_5.Text = Convert.ToString(dt14.Rows[0]["Summary"]);


                DataTable dt15 = ds.Tables[15];
                if (dt15 != null)
                {
                    if (dt15.Rows.Count > 0)
                    {
                        lblSectionTitleOnline_5.Text = Convert.ToString(dt15.Rows[0]["SectionTitle"]);
                        lvOnlineArticles_5.DataSource = dt15;
                        lvOnlineArticles_5.DataBind();
                    }
                }

                DataTable dt16 = ds.Tables[16];
                if (dt16 != null)
                {
                    if (dt16.Rows.Count > 0)
                    {
                        lblSectionTitlePrint_5.Text = Convert.ToString(dt16.Rows[0]["SectionTitle"]);
                        lvPrintArticles_5.DataSource = dt16;
                        lvPrintArticles_5.DataBind();
                    }
                }
            }


           
        }
    }
}