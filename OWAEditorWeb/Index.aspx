<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="OWAEditorWeb.Index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:TextBox runat="server" ID="txtId"></asp:TextBox>
        <asp:Button runat="server" ID="btnE" OnClick="btnE_Click"  Text="签出"/>
        <asp:Button runat="server" ID="btnB" OnClick="btnB_Click"  Text="编辑"/>
        <asp:Button runat="server" ID="btnS" OnClick="btnS_Click"  Text="签入"/>
        <asp:Button runat="server" ID="btnC" OnClick="btnC_Click"  Text="清缓存"/>
        <br />

        <asp:Literal runat="server" ID="litList"></asp:Literal>
    </div>
    </form>
</body>
</html>
