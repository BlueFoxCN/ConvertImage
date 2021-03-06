﻿using System;
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

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ConvertImage
{
    public partial class Generate : System.Web.UI.Page
    {
        private int[] LEN_THRESH = new int[] { 10, 20 };
        private int LINE_LEN = 80;
        private string image_path = "";
        private WebClient wc = new WebClient();
        private string[] item_ary = new string[] { "A", "B", "C", "D" };
        protected void Page_Load(object sender, EventArgs e)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(typeof(Generate));
            try
            {
                Document doc = new Document(@Server.MapPath("templates") + "\\" + "template.docx");
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToDocumentEnd();


                string dataStr = Request.Params["data"];
                if (dataStr == null || dataStr.Trim() == string.Empty)
                    return;
                dataStr = dataStr.Replace("=>", ":");

                var serializer = new JavaScriptSerializer();
                serializer.RegisterConverters(new[] { new DynamicJsonConverter() });
                dynamic data = serializer.Deserialize<object>(dataStr);

                log.Info(data);

                ArrayList questions = new ArrayList(data.questions);
                string fileName = data.name;
                string qrcodeHost = data.qrcode_host;
                string docType = data.doc_type;
                bool qr_code = data.qr_code;
                bool withNumber = data.with_number;
                bool firstLine = false;
                bool withAnswer = data.with_answer;
                int number = 1;
                string para_with_number = "";
                

                Shape shape;
                string qrcode_path = "";

                if (data.app_qr_code)
                {
                    // should insert app download info at the beginning
                    string android_app_qrcode_path = @Server.MapPath("public\\qrcodes\\android_app_qrcode.png");
                    string ios_app_qrcode_path = @Server.MapPath("public\\qrcodes\\ios_app_qrcode.png");
                    if (!File.Exists(android_app_qrcode_path))
                    {
                        HttpWebRequest httpRequest = (HttpWebRequest)
                            WebRequest.Create(qrcodeHost + "/student_android_app_qr_code");
                        httpRequest.Method = WebRequestMethods.Http.Get;

                        HttpWebResponse httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                        using (Stream inputStream = httpResponse.GetResponseStream())
                        using (Stream outputStream = File.OpenWrite(android_app_qrcode_path))
                        {
                            byte[] buffer = new byte[4096];
                            int bytesRead;
                            do
                            {
                                bytesRead = inputStream.Read(buffer, 0, buffer.Length);
                                outputStream.Write(buffer, 0, bytesRead);
                            } while (bytesRead != 0);
                        }
                    }
                    if (!File.Exists(ios_app_qrcode_path))
                    {
                        HttpWebRequest httpRequest = (HttpWebRequest)
                            WebRequest.Create(qrcodeHost + "/student_ios_app_qr_code");
                        httpRequest.Method = WebRequestMethods.Http.Get;

                        HttpWebResponse httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                        using (Stream inputStream = httpResponse.GetResponseStream())
                        using (Stream outputStream = File.OpenWrite(ios_app_qrcode_path))
                        {
                            byte[] buffer = new byte[4096];
                            int bytesRead;
                            do
                            {
                                bytesRead = inputStream.Read(buffer, 0, buffer.Length);
                                outputStream.Write(buffer, 0, bytesRead);
                            } while (bytesRead != 0);
                        }
                    }

                    builder.Write("Android客户端下载：");
                    shape = builder.InsertImage(android_app_qrcode_path);
                    shape.WrapType = WrapType.Inline;
                    shape.VerticalAlignment = VerticalAlignment.Inline;
                    builder.Write("     ");
                    builder.Write("iOS客户端下载：");
                    shape = builder.InsertImage(ios_app_qrcode_path);
                    shape.WrapType = WrapType.Inline;
                    shape.VerticalAlignment = VerticalAlignment.Inline;
                    builder.Write("\n");
                    builder.Writeln("学生网页版登录地址： " + data.student_portal_url);
                    builder.Writeln("");
                }

                foreach (dynamic q in questions)
                {
                    if (q.image_path == null)
                    {
                        image_path = "";
                    }
                    else
                    {
                        image_path = q.image_path;
                    }
                    if (q.link != null && qr_code)
                    {
                        // first get the qr_code image
                        qrcode_path = @Server.MapPath("public\\qrcodes\\" + q.link + ".png");
                        if (!File.Exists(qrcode_path))
                        {
                            HttpWebRequest httpRequest = (HttpWebRequest)
                                WebRequest.Create(qrcodeHost + "/qrcodes?link=" + q.link);
                            httpRequest.Method = WebRequestMethods.Http.Get;

                            HttpWebResponse httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                            using (Stream inputStream = httpResponse.GetResponseStream())
                            using (Stream outputStream = File.OpenWrite(qrcode_path))
                            {
                                byte[] buffer = new byte[4096];
                                int bytesRead;
                                do
                                {
                                    bytesRead = inputStream.Read(buffer, 0, buffer.Length);
                                    outputStream.Write(buffer, 0, bytesRead);
                                } while (bytesRead != 0);
                            }
                        }

                        // shape = builder.InsertImage(httpResponseStream);
                        shape = builder.InsertImage(qrcode_path);
                        shape.WrapType = WrapType.Square;
                        shape.Left = 370;
                    }

                    // insert the question content
                    firstLine = true;
                    foreach (dynamic para in q.content)
                    {
                        if (para is string)
                        {
                            if (withNumber && firstLine)
                            {
                                firstLine = false;
                                para_with_number = number.ToString() + ". " + para;
                                number++;
                            }
                            else
                            {
                                para_with_number = para;
                            }
                            writeParagraph(builder, (string)para_with_number);
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

                    // insert answers
                    if (withAnswer)
                    {
                        if (q.answer != -1 || q.answer_content.Count > 0)
                        {
                            Font font = builder.Font;
                            font.Bold = true;
                            builder.Write("解答：");
                            font.Bold = false;
                            if (q.answer != -1)
                            {
                                builder.Write(item_ary[q.answer]);
                            }
                            builder.Writeln("");
                            foreach (dynamic para in q.answer_content)
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
                        }
                    }


                    for (int i = 0; i < 3; i++)
                    {
                        writeParagraph(builder, "");
                    }
                }
                string finalName = "public/documents/" + fileName + "-" + Guid.NewGuid().ToString();
                finalName += docType == "word" ? ".docx" : ".pdf";
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
                    if (image_path == "")
                    {
                        shape = builder.InsertImage(@Server.MapPath("public\\download\\" + imageInfo[0] + "." + imageInfo[1]));
                    }
                    else
                    {
                        shape = builder.InsertImage(wc.DownloadData(image_path + imageInfo[0] + "." + imageInfo[1]));
                    }
                    shape.Width = Convert.ToDouble(imageInfo[2]);
                    shape.Height = Convert.ToDouble(imageInfo[3]);
                    shape.WrapType = WrapType.Inline;
                    shape.VerticalAlignment = VerticalAlignment.Inline;
                }
                else if (ele.StartsWith("fig_"))
                {
                    imageInfo = ele.Substring(4).Split('*');
                    if (image_path == "")
                    {
                        shape = builder.InsertImage(@Server.MapPath("public\\download\\" + imageInfo[0] + "." + imageInfo[1]));
                    }
                    else
                    {
                        shape = builder.InsertImage(wc.DownloadData(image_path + imageInfo[0] + "." + imageInfo[1]));
                    }
                    shape.Width = Convert.ToDouble(imageInfo[2]);
                    shape.Height = Convert.ToDouble(imageInfo[3]);
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