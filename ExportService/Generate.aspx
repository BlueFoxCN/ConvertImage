<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Generate.aspx.cs" Inherits="ConvertImage.Generate" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <br />
        <asp:TextBox ID="data" runat="server" Height="100px" 
            TextMode="MultiLine" Width="400px"></asp:TextBox>
        <br />
        <br />
        <input type="submit" id="Submit1" value="Upload" runat="server" />
    
    </div>
    </form>
</body>
</html>
