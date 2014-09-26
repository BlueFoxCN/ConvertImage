using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Collections;
using ImageMagick;
using System.Web.Script.Serialization;

namespace ConvertImage
{
    public partial class WebForm1 : System.Web.UI.Page
    {

        public string path = HttpContext.Current.Server.MapPath("~/");
        public int minFigHeight = 50;
        protected void Page_Load(object sender, EventArgs e)
        {
            string SaveLocation = "";
            if ((file.PostedFile != null) && (file.PostedFile.ContentLength > 0))
            {
                string fn = System.IO.Path.GetFileName(file.PostedFile.FileName);
                SaveLocation = Server.MapPath("data") + "\\" + fn;
                try
                {
                    file.PostedFile.SaveAs(SaveLocation);
                }
                catch (Exception ex)
                {
                    Response.Clear();
                    Response.Write("Error: " + ex.Message);
                }
            }
            else
            {
                return;
            }
            
            Document doc = new Document(@SaveLocation);

            ArrayList content = new ArrayList();

            foreach (Node node in doc.Sections[0].Body.GetChildNodes(NodeType.Any, false)) {
                switch (node.NodeType)
                {
                    case NodeType.Paragraph:
                        content.AddRange(parseParagraph((Paragraph)node));
                        break;
                    case NodeType.Table:
                        content.Add(parseTable((Aspose.Words.Tables.Table)node));
                        break;
                    case NodeType.Shape:
                        content.Add(node);
                        break;
                }
            }

            Response.Clear();
            var json = new JavaScriptSerializer().Serialize(content);
            Response.Write(json);

            Response.End();
        }

        public ArrayList parseParagraph(Paragraph p)
        {
            ArrayList content = new ArrayList();
            string curText = "";
            string suffix = "";
            string imgFileName = "";
            bool skip = false;
            foreach (Node node in p.GetChildNodes(NodeType.Any, false)) {
                
                if (node.NodeType == NodeType.FieldStart)
                {
                    skip = true;
                }
                else if (node.NodeType == NodeType.FieldSeparator || node.NodeType == NodeType.FieldEnd)
                {
                    skip = false;
                    continue;
                }
                if (skip == true)
                {
                    continue;
                }
                
                string[] typeInfo = judgeType(node);
                switch (typeInfo[0])
                {
                    case "text":
                        if (node.NodeType == NodeType.SmartTag)
                        {
                            curText += node.GetText();
                        }
                        else
                        {
                            if (((Run)node).Font.Subscript)
                            {
                                curText += "$$sub_" + node.GetText() + "$$";
                            }
                            else if (((Run)node).Font.Superscript)
                            {
                                curText += "$$sup_" + node.GetText() + "$$";
                            }
                            else if (((Run)node).Font.Underline == Underline.Single)
                            {
                                curText += "$$und_" + node.GetText() + "$$";
                            }
                            else if (((Run)node).Font.Italic)
                            {
                                curText += "$$ita_" + node.GetText() + "$$";
                            }
                            else if (node.GetText() == "\v")
                            {
                                content.Add(curText);
                                curText = "";
                            }
                            else
                            {
                                curText += node.GetText();
                            }
                        }
                        break;
                    case "unknown":
                        curText += node.GetText();
                        break;
                    case "equation":
                        suffix = getSuffix(node);
                        if (suffix == "noimage")
                            continue;
                        imgFileName = Guid.NewGuid().ToString();
                        convertImage(node, imgFileName);
                        curText += "$$equ_" + imgFileName + "*" + typeInfo[1] + "*" + typeInfo[2] + "$$";
                        break;
                    case "figure":
                        if (curText != "")
                            content.Add(curText);
                        curText = "";
                        suffix = getSuffix(node);
                        if (suffix == "noimage")
                            continue;
                        imgFileName = Guid.NewGuid().ToString();
                        convertImage(node, imgFileName);
                        content.Add("$$fig_" + imgFileName + "*" + typeInfo[1] + "*" + typeInfo[2] + "$$");
                        break;
                }
            }
            if (curText != "")
                content.Add(curText);
            if (content.Count == 0)
                content.Add("");
            return content;
        }

        public Hashtable parseTable(Aspose.Words.Tables.Table t)
        {
            ArrayList curRow = new ArrayList();
            ArrayList curCell = new ArrayList();
            Hashtable h = new Hashtable();
            h.Add("type", "table");
            ArrayList content = new ArrayList();
            foreach (Aspose.Words.Tables.Row row in t.GetChildNodes(NodeType.Row, false))
            {
                curRow.Clear();
                foreach (Aspose.Words.Tables.Cell cell in row.GetChildNodes(NodeType.Cell, false))
                {
                    curCell.Clear();
                    foreach (Paragraph node in cell.GetChildNodes(NodeType.Paragraph, false))
                    {
                        curCell.AddRange(parseParagraph(node));
                    }
                    curRow.Add(curCell.ToArray());
                }
                content.Add(curRow.ToArray());
            }
            h.Add("content", content.ToArray());
            return h;
        }

        public string[] judgeType(Node node)
        {
            if (node.NodeType == NodeType.Run || node.NodeType == NodeType.SmartTag)
            {
                return new string[] {"text", "", ""};
            }
            else if (node.NodeType == NodeType.Shape && ((Shape)node).ImageData.ImageType == ImageType.Wmf)
            {
                return new string[] { "equation", ((Shape)node).Width.ToString(), ((Shape)node).Height.ToString() };
            }
            else if (node.NodeType == NodeType.Shape && !((Shape)node).IsInline)
            {
                return new string[] { "figure", ((Shape)node).Width.ToString(), ((Shape)node).Height.ToString() };
            }
            else if (node.NodeType == NodeType.DrawingML && ((DrawingML)node).Height < minFigHeight)
            {
                return new string[] { "equation", ((DrawingML)node).Width.ToString(), ((DrawingML)node).Height.ToString() };
            }
            else if (node.NodeType == NodeType.DrawingML)
            {
                return new string[] { "figure", ((DrawingML)node).Width.ToString(), ((DrawingML)node).Height.ToString() };
            }
            else if (node.NodeType == NodeType.Shape && ((Shape)node).IsInline)
            {
                return new string[] { "equation", ((Shape)node).Width.ToString(), ((Shape)node).Height.ToString() };
            }
            else
            {
                return new string[] { "unknown", "", "" };
            }
        }

        public string getSuffix(Node node)
        {
            if (node.NodeType == NodeType.DrawingML)
            {
                return ((DrawingML)node).ImageData.ImageType.ToString().ToLower();
            }
            else if (node.NodeType == NodeType.Shape)
            {
                return ((Shape)node).ImageData.ImageType.ToString().ToLower();
            }
            else
            {
                return "";
            }
        }

        public void convertImage(Node node, string filename)
        {
            byte[] id;
            if (node.NodeType == NodeType.Shape)
            {
                id = ((Shape)node).ImageData.ToByteArray();
            }
            else if (node.NodeType == NodeType.DrawingML)
            {
                id = ((DrawingML)node).ImageData.ToByteArray();
            }
            else
            {
                return;
            }
            using (MagickImage image = new MagickImage(id))
            {
                image.Transparent(new MagickColor("#FFFFFFFF"));
                if (image.Format == MagickFormat.Wmf)
                {
                    image.Resize(new Percentage(0.35));
                }
                image.Write(path + "public\\download\\" + filename + ".png");
            }
        }
    }
}