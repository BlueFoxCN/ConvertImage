using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.Script.Serialization;
using System.Web.UI.WebControls;
using Aspose.Words;
using Aspose.Words.Drawing;
using ImageMagick;
using System.Collections;
using System.Net;
using System.IO;

namespace ConvertImage
{
    public partial class ExportNote : System.Web.UI.Page
    {
        private int[] LEN_THRESH = new int[] { 10, 20 };
        private int LINE_LEN = 80;
        protected void Page_Load(object sender, EventArgs e)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(typeof(Generate));
            try
            {
                Document doc = new Document(@Server.MapPath("templates") + "\\" + "template.docx");
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToDocumentEnd();


                string dataStr = Request.Params["notes"];
                if (dataStr == null || dataStr.Trim() == string.Empty)
                    return;
                dataStr = dataStr.Replace("=>", ":");

                var serializer = new JavaScriptSerializer();
                serializer.RegisterConverters(new[] { new DynamicJsonConverter() });
                dynamic data = serializer.Deserialize<object>(dataStr);

                log.Info(data);

                ArrayList notes = new ArrayList(data);
                string fileName = "错题本";

                foreach (dynamic q in notes)
                {
                    // insert the question content
                    foreach (dynamic para in q.content)
                    {
                        if (para is string)
                        {
                            writeParagraph(builder, (string)para);
                        }
                        else
                        {
                            switch ((string)(para["type"]))
                            {
                                case "table":
                                    writeTable(builder, para["content"]);
                                    break;
                            }
                        }
                    }

                    // insert the items if this is a choice question
                    if (q.type == "choice")
                    {
                        List<string> itemLines = organizeItems(q.items);
                        foreach (string itemLine in itemLines)
                        {
                            writeParagraph(builder, itemLine);
                        }
                    }

                    // insert note if there is any
                    if (q.note_type != null && q.note_type != 0)
                    {
                        writeNoteType(builder, q.note_type);
                    }

                    if (q.topics != null && q.topics.Count > 0)
                    {
                        writeTopics(builder, q.topics);
                    }

                    if (q.summary != null && q.summary != "")
                    {
                        writeSummary(builder, q.summary);
                    }

                    for (int i = 0; i < 3; i++)
                    {
                        writeParagraph(builder, "");
                    }
                }
                string finalName = "public/documents/" + fileName + "-" + Guid.NewGuid().ToString() + ".docx";
                doc.Save(@Server.MapPath(finalName));
                Response.Clear();
                Response.Write(finalName);
            }
            catch (Exception es)
            {
                Response.Clear();
                Response.Write("error: " + es);
            }
            Response.End();
        }

        private void writeSummary(DocumentBuilder builder, string summary)
        {
            Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.Red;
            builder.Write("总结：" + summary);
            builder.Writeln("");
            font.Bold = false;
            font.Color = System.Drawing.Color.Black;
        }

        private void writeTopics(DocumentBuilder builder, List<object> topicsList)
        {
            string[] topicsAry = new string[topicsList.Count];
            for (int i = 0; i < topicsList.Count; i++)
            {
                topicsAry[i] = topicsList[i].ToString();
            }
            string topicsStr = String.Join("，", topicsAry);
            if (topicsStr == "")
            {
                return;
            }

            Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.Red;
            builder.Write("知识点：" + topicsStr);
            builder.Writeln("");
            font.Bold = false;
            font.Color = System.Drawing.Color.Black;
        }

        private void writeNoteType(DocumentBuilder builder, int noteType)
        {
            Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.Red;
            builder.Write("错误类型：");
            switch (noteType)
            {
                case 1:
                    builder.Write("不懂");
                    break;
                case 2:
                    builder.Write("不会");
                    break;
                case 3:
                    builder.Write("不对");
                    break;
            }
            builder.Writeln("");
            font.Bold = false;
            font.Color = System.Drawing.Color.Black;
        }

        private List<string> organizeItems(List<object> rawItems)
        {
            List<string> items = new List<string>();
            List<string> itemLines = new List<string>();
            int maxLen = 0;
            string ele;
            for (int i = 0; i < rawItems.Count; i++)
            {
                ele = decimalToCapital(i) + ". " + rawItems[i];
                items.Add(ele);
                if (ele.Length > maxLen)
                    maxLen = ele.Length;
            }

            if (maxLen < LEN_THRESH[0])
            {
                itemLines.Add(items[0].PadRight(LINE_LEN / 4 - items[0].Length) +
                    items[1].PadRight(LINE_LEN / 4 - items[1].Length) +
                    items[2].PadRight(LINE_LEN / 4 - items[2].Length) +
                    items[3]);
            }
            else if (maxLen < LEN_THRESH[1])
            {
                itemLines.Add(items[0].PadRight(LINE_LEN / 2 - items[0].Length) +
                    items[1]);
                itemLines.Add(items[2].PadRight(LINE_LEN / 2 - items[2].Length) +
                    items[3]);
            }
            else
            {
                return items;
            }
            return itemLines;
        }

        private string decimalToCapital(int d)
        {
            string[] c = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" };
            return c[d];
        }

        private void writeFigure(DocumentBuilder builder, string fig)
        {
            string[] imageInfo = fig.Substring(4).Split('*');
            Shape shape = builder.InsertImage("public\\download\\" + imageInfo[0] + ".png");
            shape.Width = Convert.ToDouble(imageInfo[1]);
            shape.Height = Convert.ToDouble(imageInfo[2]);
            shape.WrapType = WrapType.Inline;
        }

        private void writeParagraph(DocumentBuilder builder, string p, bool newLine = true)
        {
            Font font = builder.Font;
            string[] content = p.Split(new string[] { "$$" }, StringSplitOptions.None);
            string[] imageInfo;
            Shape shape;
            foreach (string ele in content)
            {
                if (ele.StartsWith("math_"))
                {
                    imageInfo = ele.Substring(5).Split('*');
                    Document mathDoc = new Document(@Server.MapPath("public\\mathtype\\" + imageInfo[0] + ".docx"));
                    Paragraph mathP = (Paragraph)mathDoc.Sections[0].Body.ChildNodes[0];
                    Shape mathShape = (Shape)mathP.ChildNodes[2];

                    Shape newShape = (Shape)builder.Document.ImportNode(mathShape, true);
                    builder.InsertNode(newShape);
                }
                else if (ele.StartsWith("equ_"))
                {
                    imageInfo = ele.Substring(4).Split('*');
                    shape = builder.InsertImage(@Server.MapPath("public\\download\\" + imageInfo[0] + ".png"));
                    shape.Width = Convert.ToDouble(imageInfo[1]);
                    shape.Height = Convert.ToDouble(imageInfo[2]);
                    shape.WrapType = WrapType.Inline;
                    shape.VerticalAlignment = VerticalAlignment.Inline;
                }
                else if (ele.StartsWith("fig_"))
                {
                    imageInfo = ele.Substring(4).Split('*');
                    shape = builder.InsertImage(@Server.MapPath("public\\download\\" + imageInfo[0] + ".png"));
                    shape.Width = Convert.ToDouble(imageInfo[1]);
                    shape.Height = Convert.ToDouble(imageInfo[2]);
                    shape.WrapType = WrapType.Inline;
                    shape.VerticalAlignment = VerticalAlignment.Inline;
                }
                else if (ele.StartsWith("sub_"))
                {
                    font.Subscript = true;
                    builder.Write(ele.Substring(4));
                    font.Subscript = false;
                }
                else if (ele.StartsWith("sup_"))
                {
                    font.Superscript = true;
                    builder.Write(ele.Substring(4));
                    font.Superscript = false;
                }
                else if (ele.StartsWith("und_"))
                {
                    font.Underline = Underline.Single;
                    builder.Write(ele.Substring(4));
                    font.Underline = Underline.None;
                }
                else if (ele.StartsWith("ita_"))
                {
                    font.Italic = true;
                    builder.Write(ele.Substring(4));
                    font.Italic = false;
                }
                else
                {
                    builder.Write(ele);
                }
            }
            if (newLine)
                builder.Writeln("");
        }

        private void writeTable(DocumentBuilder builder, ArrayList content)
        {
            builder.StartTable();
            foreach (ArrayList row in content)
            {
                foreach (ArrayList cell in row)
                {
                    builder.InsertCell();
                    foreach (string para in cell)
                    {
                        writeParagraph(builder, para, false);
                    }
                }
                builder.EndRow();
            }
            builder.EndTable();
            builder.Writeln("");
        }
    }
}