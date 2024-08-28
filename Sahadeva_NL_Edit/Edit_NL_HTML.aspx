<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Edit_NL_HTML.aspx.cs" Inherits="Sahadeva_NL_Edit.Edit_NL_HTML" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!DOCTYPE html>
<html >
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    
    <title>AdFactorsPR Newsletter Template</title>
    <style type="text/css">
    @import url("https://fonts.googleapis.com/css2?family=Lato:wght@700;900&family=Noto+Sans+Georgian&family=Noto+Serif+Georgian:wght@400;700&family=Open+Sans:wght@400;600&family=Roboto:wght@100&display=swap");        
      body {
        margin: 0;
        background-color: #f0f0f0;
      }
      table {
        border-spacing: 0;
      }
      td {
        padding: 0;
      }
      img {
        border: 0;
      }
      .wrapper {
        width: 100%;
        table-layout: fixed;
        background-color: #f0f0f0;
        padding: 40px 0px;
      }
      .main {
        background-color: #fff;
        margin: 0 auto;
        width: 100%;
        /*max-width: 600px;*/
        border-spacing: 0;
        border-radius: 10px;
        font-family: "Noto Sans Georgian", sans-serif;
        color: #393939;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
      }
     .customTable {
        width: 100%;
        background-color: #F2F8FB;
        /* border-collapse: collapse; */
        border-width: 1px;
        border-color: #eaebec;
        border-style: solid;
        color: #000000;
        margin: 20px 0px;
        border-radius: 4px;
        overflow: hidden;
      }
        .customTable tr.border_bottom td
        {
            border-bottom: 1px solid #DAE4E4;
        }
        
        
        
            .card-title
        {
            margin: -10px;
            font-size: 15px;
        }
        .card-body
        {
            padding-top: 0.45rem;
            padding-right: 1.25rem;
            padding-bottom: 0rem;
            padding-left: 1.25rem;
        }
        .form-group
        {
            margin-bottom: 0.5rem;
        }
        .dropdown-menu
        {
            font-size: 0.8rem;
        }
        input
        {
            font-size: .75rem !important;
        }
        .textarea
        {
            font-size: .75rem !important;
        }
        .form-control
        {
            font-size: .75rem !important;
            border: 1px solid #aaa;
        }
        .multiselect-selected-text
        {
            font-size: .75rem !important;
        }
        .multiselect
        {
            border: 1px solid #aaa;
        }
        
        .dataTables_filter input
        {
            margin-left: 0px;
        }
        
        
        fieldset.scheduler-border
        {
            border: 1px groove #ddd !important;
            padding: 0 1.4em 1.4em 1.4em !important;
            margin: 0 0 1.5em 0 !important;
            -webkit-box-shadow: 0px 0px 0px 0px #000;
            box-shadow: 0px 0px 0px 0px #000;
            border-top: 1px groove #ddd !important;
            border-left: 1px groove #ddd !important;
            border-right: 1px groove #ddd !important;
        }
        
        legend.scheduler-border
        {
            font-size: 1.2em !important;
            font-weight: bold !important;
            text-align: left !important;
            width: auto;
            padding: 0 10px;
            border-bottom: none;
            color: mediumslateblue;
        }
        
        
        .table th, .table td {
            padding: .50rem;
            vertical-align: top;
            border: 1px solid #7f848a;
        }
        
        .table thead th {
            vertical-align: bottom;
            border: 1px solid #7f848a;
        }
        
    
    </style>
    
    
</head>
<body>
    <center class="wrapper">
        <table class="main" width="100%" style="padding: 40px">
            <!-- LOGO SECTION and Title info -->
            <tr>
                <td>
                    <img src="#EntityLogoUrl#" alt="logo" />
                    <div style="font-size: 40px; color: #373469; font-weight: bolder; padding-bottom: 20px;">
                        <asp:Label ID="lblEntityName" runat="server">EntityName</asp:Label>
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%">
                        <tr>
                            <td class="two-columns"  style="text-align: right">
                                <div style="display: flex; justify-content: space-between; align-items: center;">
                                    <div style="border-bottom: 2px solid #41a499; padding-bottom: 4px; font-family: 'Lato', sans-serif;
                                        font-size: 20px; font-weight: 900; letter-spacing: 1px; text-align: center; color: #333;">
                                        <asp:Label ID="lblNewsLetterTitle" runat="server">NewsLetterTitle</asp:Label>                                        
                                    </div>
                                   
                                </div>
								 <div style="font-size: 14px; color: #41a499; font-weight: 700; font-family: 'Lato', sans-serif;">
                                        <asp:Label ID="lblNewsLetterDate" runat="server">NewsLetterDate</asp:Label>
                                 </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <!-- Title info section  -->
            
           <!--  Client 1  -->

            <tr>
            <td>
             <div style="padding: 20px 0px">
                <div style="padding: 20px 0px">
                <div style="padding-left: 10px;
                font-family: 'Open Sans', sans-serif;
                font-size: 18px;
                font-weight: bold;
                letter-spacing: 0.36px;
                text-align: left;
                color: #393939;">
                <asp:Label ID="lblEntityNameSection_1" runat="server">EntityName</asp:Label>
                </div>
                <div>
                <ul style="padding: 0px; font-size: 12px; font-family: Georgia; font-size: 12px;
                text-align: justify; color: #393939;">
                <asp:Label ID="lblEntityDescription_1" runat="server">EntityDescription</asp:Label>                
                </ul>
                </div>
                </div>
                </div>
            </td>
</tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitleOnline_1" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvOnlineArticles_1" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 20%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                            <%--<th style="width: 5%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Action
                                            </th>--%>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                   <%-- <td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("ArticleID"),Eval("MediaTypeID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />                                        
                                        <asp:Button ID="btnUpdate" runat="server" Text="Update Question File" CommandArgument='<%# Eval("ArticleID"),Eval("MediaTypeID")%>'
                                            CommandName="UpdateDetails" CssClass="formbtn" />
                                        <asp:Button ID="btnDelete" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("ArticleID"),Eval("MediaTypeID")%>'
                                            CommandName="DeleteRecord" CssClass="formbtn" />
                                        
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitlePrint_1" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvPrintArticles_1" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                     <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

        <!--  Client 2  -->

            <tr>
            <td>
             <div style="padding: 20px 0px">
                <div style="padding: 20px 0px">
                <div style="padding-left: 10px;
                font-family: 'Open Sans', sans-serif;
                font-size: 18px;
                font-weight: bold;
                letter-spacing: 0.36px;
                text-align: left;
                color: #393939;">
                <asp:Label ID="lblEntityNameSection_2" runat="server">EntityName</asp:Label>
                </div>
                <div>
                <ul style="padding: 0px; font-size: 12px; font-family: Georgia; font-size: 12px;
                text-align: justify; color: #393939;">
                <asp:Label ID="lblEntityDescription_2" runat="server">EntityDescription</asp:Label>                
                </ul>
                </div>
                </div>
                </div>
            </td>
</tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitleOnline_2" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvOnlineArticles_2" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitlePrint_2" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvPrintArticles_2" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

        <!--  Client 3  -->

            <tr>
            <td>
             <div style="padding: 20px 0px">
                <div style="padding: 20px 0px">
                <div style="padding-left: 10px;
                font-family: 'Open Sans', sans-serif;
                font-size: 18px;
                font-weight: bold;
                letter-spacing: 0.36px;
                text-align: left;
                color: #393939;">
                <asp:Label ID="lblEntityNameSection_3" runat="server">EntityName</asp:Label>
                </div>
                <div>
                <ul style="padding: 0px; font-size: 12px; font-family: Georgia; font-size: 12px;
                text-align: justify; color: #393939;">
                <asp:Label ID="lblEntityDescription_3" runat="server">EntityDescription</asp:Label>                
                </ul>
                </div>
                </div>
                </div>
            </td>
</tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitleOnline_3" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvOnlineArticles_3" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitlePrint_3" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvPrintArticles_3" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

        <!--  Client 4  -->

            <tr>
            <td>
             <div style="padding: 20px 0px">
                <div style="padding: 20px 0px">
                <div style="padding-left: 10px;
                font-family: 'Open Sans', sans-serif;
                font-size: 18px;
                font-weight: bold;
                letter-spacing: 0.36px;
                text-align: left;
                color: #393939;">
                <asp:Label ID="lblEntityNameSection_4" runat="server">EntityName</asp:Label>
                </div>
                <div>
                <ul style="padding: 0px; font-size: 12px; font-family: Georgia; font-size: 12px;
                text-align: justify; color: #393939;">
                <asp:Label ID="lblEntityDescription_4" runat="server">EntityDescription</asp:Label>                
                </ul>
                </div>
                </div>
                </div>
            </td>
</tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitleOnline_4" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvOnlineArticles_4" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitlePrint_4" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvPrintArticles_4" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>


        <!--  Client 5  -->

            <tr>
            <td>
             <div style="padding: 20px 0px">
                <div style="padding: 20px 0px">
                <div style="padding-left: 10px;
                font-family: 'Open Sans', sans-serif;
                font-size: 18px;
                font-weight: bold;
                letter-spacing: 0.36px;
                text-align: left;
                color: #393939;">
                <asp:Label ID="lblEntityNameSection_5" runat="server">EntityName</asp:Label>
                </div>
                <div>
                <ul style="padding: 0px; font-size: 12px; font-family: Georgia; font-size: 12px;
                text-align: justify; color: #393939;">
                <asp:Label ID="lblEntityDescription_5" runat="server">EntityDescription</asp:Label>                
                </ul>
                </div>
                </div>
                </div>
            </td>
</tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitleOnline_5" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvOnlineArticles_5" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

            <tr>
          <td>
            <div
              style="
                padding: 0px 14px;
                border-left: 4px solid #018577;
                font-family: 'Open Sans', sans-serif;
                font-size: 16px;
                font-weight: 800;
                font-stretch: normal;
                font-style: normal;
                line-height: normal;
                letter-spacing: 0.32px;
                text-align: left;
                color: #393939;
              ">
                <asp:Label ID="lblSectionTitlePrint_5" runat="server"></asp:Label>       
            </div>
          </td>
        </tr>
		       
            <tr>
          <td>
            <div class="container-fluid">
               
                <!-- /.row -->
                <div class="row mb-2">
                    <div class="col-sm-12">
                        <asp:ListView ID="lvPrintArticles_5" runat="server" Visible="true">
                            <LayoutTemplate>
                                <table class="customTable" id="TestTable" style="width: 100%">
                                    <thead>
                                        <tr style="border: 1px solid #969ba1">
                                             <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                ArticleID
                                            </th>
                                            <th style="width: 2%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                MediaTypeID
                                            </th>
                                            <th style="width: 13%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Cluster
                                            </th>
                                            <th style="width: 25%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Headline
                                            </th>
                                            <th style="width: 30%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Summary
                                            </th>
                                            <th style="width: 8%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Sentiment
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Publication
                                            </th>
                                            <th style="width: 10%; text-align: center; border-bottom: 1px solid #eaebec;padding: 10px;font-family: Georgia;font-size: 12px;text-align: left;"">
                                                Article Type
                                            </th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr id="itemPlaceholder" runat="server" />
                                    </tbody>
                                </table>
                                <%--</div>--%>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <tr style="background-color: #EBF6F5 ;" class="border_bottom">
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("MediaTypeID")%> &nbsp;
                                        <%--<%# Eval("Sr_No") %>--%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <div style="background-color: #ClusterColour#; border-radius: 3px; color: #000000; width: fit-content;padding: 4px 4px;font-weight:bold;">
                                         <%# Eval("Cluster")%>
                                         </div>                                          
                                    </td >
                                    <td style="text-align: center; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;" >
                                        <%# Eval("Headline")%>
                                    </td>
                                    <td  style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Summary")%>
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">

                                        <div
                                          style="
                                            display: flex;
                                            justify-content: center;
                                            gap: 10px;
                                          "
                                        >
                                          <%# Eval("Sentiment")%> &nbsp;&nbsp;
                                          <div
                                            style="
                                              margin: 2px 0px;
                                              background-color: <%# Eval("SentimentColor")%>;
                                              width: 10px;
                                              height: 10px;
                                              border-radius: 50%;
                                            "
                                          ></div>
                                        </div>                                        
                                    </td>
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("Publication")%>
                                    </td>
                                    
                                    <td style="text-align: left; padding: 15px 12px;font-family: Georgia;font-size: 12px;text-align: left;color: #393939;vertical-align: top;">
                                        <%# Eval("ArticleType") %>
                                    </td>
                                    <%--<td>
                                        <asp:Button ID="btnEdit" runat="server" Text="Edit" CommandArgument='<%# Eval("JOBID")%>'
                                            OnClientClick="target ='_blank';" CommandName="Edit" CssClass="formbtn" />
                                        <asp:Button ID="btnDownloadOutput" runat="server" Text="View" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="DownloadOutput" CssClass="formbtn" />
                                        <asp:Button ID="btnJobProgress" runat="server" Text="Check Job Progress" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="CheckJobProgress" CssClass="formbtn" />
                                        <asp:Button ID="btnQuestion" runat="server" Text="Update Question File" CommandArgument='<%# Eval("JOBID")%>'
                                            CommandName="UpdateQuestionFile" CssClass="formbtn" />
                                    </td>--%>
                                </tr>
                            </ItemTemplate>
                            <EmptyDataTemplate>
                            </EmptyDataTemplate>
                        </asp:ListView>
                    </div>
                    <!-- /.col -->
                </div>
                <!-- /.row -->
            </div>
          </td>
        </tr>

            <!-- FOOTER SECTION -->
            <tr>
                <td align="center">
                    <img src="http://203.212.222.191:9553/Areas/Admin/image/logo_small.png" alt="" />
                    </td>
            </tr>
            <tr>
                <td align="center">
                    <div style="padding: 10px; font-family: Georgia; font-size: 13px; text-align: center;
                        color: #41414e;">
                        All copyrights reserved
                    </div>
                </td>
            </tr>
        </table>
    </center>
</body>
</html>
