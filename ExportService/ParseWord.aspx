<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ParseWord.aspx.cs" Inherits="ConvertImage.ParseWord" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" method="post" enctype="multipart/form-data" runat="server">
    <INPUT type=file id=word_file name=word_file runat="server" />
    <INPUT type=text id=job_id name=job_id runat="server" />
    <br>
    <input type="submit" id="Submit1" value="Upload" runat="server" />
    </form>
</body>
</html>
