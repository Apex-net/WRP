<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="RichPreview.aspx.cs" Inherits="WebReportPreview.RichPreview" %>

<%@ Register Assembly="CrystalDecisions.Web, Version=13.0.2000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" Namespace="CrystalDecisions.Web" TagPrefix="CR" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <CR:CrystalReportViewer ID="richCrystalReportViewer" runat="server" AutoDataBind="true" HasCrystalLogo="False" HasRefreshButton="True" ToolPanelView="ParameterPanel" />
    </div>
    </form>
</body>
</html>
