<%@ Page Title=""  Language="C#"  AutoEventWireup="true" CodeBehind="EditArticleDetails.aspx.cs" Inherits="Sahadeva_NL_Edit.EditArticleDetails" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link rel="stylesheet" type="text/css" href="styles/style.css" />
    <link rel="stylesheet" type="text/css" href="styles/media-queries.css" />
    <link href='http://fonts.googleapis.com/css?family=Open+Sans:400,600' rel='stylesheet'
        type='text/css'>
    <style>
        /* width */
        ::-webkit-scrollbar
        {
            width: 10px;
        }
        
        /* Track */
        ::-webkit-scrollbar-track
        {
            background: #f1f1f1;
        }
        
        /* Handle */
        ::-webkit-scrollbar-thumb
        {
            background: #888;
        }
        
        /* Handle on hover */
        ::-webkit-scrollbar-thumb:hover
        {
            background: #555;
        }
    </style>
    <style type="text/css">
        .switch
        {
            position: relative;
            display: inline-block;
            width: 50px;
            height: 24px;
        }
        
        .switch input
        {
            opacity: 0;
        }
        
        .slider
        {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #ccc;
            -webkit-transition: .4s;
            transition: .4s;
        }
        
        .slider:before
        {
            position: absolute;
            content: "";
            height: 16px;
            width: 16px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            -webkit-transition: .4s;
            transition: .4s;
        }
        
        input:checked + .slider
        {
            background-color: #2196F3;
        }
        
        input:focus + .slider
        {
            box-shadow: 0 0 1px #2196F3;
        }
        
        input:checked + .slider:before
        {
            -webkit-transform: translateX(26px);
            -ms-transform: translateX(26px);
            transform: translateX(26px);
        }
        
        /* Rounded sliders */
        .slider.round
        {
            border-radius: 34px;
        }
        
        .slider.round:before
        {
            border-radius: 50%;
        }
    </style>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
        function toggleControl() {
            var control = document.getElementById("chkToggle");
            if (control.checked) {
                var message = "Do you wish to update notification";
                if (confirm(message)) {
                    alert('true');
                    saveData(true);
                    alert('true1Test');
                } else {
                    alert('false');
                    control.checked = false;
                    saveData(false);
                }
            } else {
                alert('not checked');
                saveData(false);
            }
        }

        function saveData(data) {
            $.ajax({
                type: "POST",
                url: "KeywordTagging.aspx/saveToggleState",
                data: "{ 'data': '" + data + "' }",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    // Handle the response
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    // Handle the error
                }
            });
        }
    </script>
    <script>
        function toggleControl1() {
            if (confirm("Do you want to save the state?")) {
                console.log("User clicked OK");
                if (confirm(message)) {
                    saveToggleState(true);
                } else {
                    control.checked = false;
                    saveToggleState(false);
                }
            } else {
                console.log("User clicked Cancel");
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="heading1" align="center">
        Edit Article Details</div>
    <table width="100%" border="0" cellspacing="0" cellpadding="3" class="csstable">
        <tr>
            <td width="30%">
                Cluster
            </td>
            <td width="70%">
                <asp:TextBox ID="txtCluster" runat="server"></asp:TextBox>
                <asp:HiddenField ID="hfArticleID" runat="server" />
            </td>
        </tr>
        <tr>
            <td width="30%">
                Headline
            </td>
            <td width="70%">
                <asp:TextBox ID="txtHeadline" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td width="30%">
                Summary
            </td>
            <td width="70%">
                <asp:TextBox ID="txtSummary" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td width="30%">
                Sentiment
            </td>
            <td width="70%">
               <%-- <asp:TextBox ID="txtSentiment" runat="server"></asp:TextBox>--%>
               <asp:DropDownList  runat="server" ID="ddlSentiment" OnSelectedIndexChanged="ddlSentiment_OnSelectedIndexChanged">
               <asp:ListItem Text="Negative" Value="Negative"></asp:ListItem>
               <asp:ListItem Text="Positive" Value="Positive"></asp:ListItem>
               <asp:ListItem Text="Neutral" Value="Neutral"></asp:ListItem>
               </asp:DropDownList>
               <asp:Label ID="lblColour" runat="server" Visible="false"></asp:Label>
            </td>
        </tr>
        <tr>
            <td width="30%">
                Publication
            </td>
            <td width="70%">
                <asp:TextBox ID="txtPublication" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td width="30%">
                Article Type
            </td>
            <td width="70%">
                <asp:TextBox ID="txtArticleType" runat="server"></asp:TextBox>
            </td>
        </tr>        
        <tr>
            <td align="center" colspan="2">
                <asp:Button ID="btnUpdate" runat="server" Text="Update" CssClass="btn" OnClick="btnUpdate_Click" />
                &nbsp;
                <%--<asp:Button ID="btnIrrelevent" runat="server" Text="Irrelevant" CssClass="btn" OnClick="btnIrrelevent_Click" />
                &nbsp;--%>
            </td>
        </tr>
    </table>
    <div style="font-size: 18px; overflow: scroll; width: 98%; height: 800px;">
        <asp:Literal ID="ltrMailBody" runat="server"></asp:Literal>
    </div>
    </form>
</body>
</html>
