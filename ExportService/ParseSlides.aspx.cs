using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Pptx;
using ImageMagick;
using System.Web.Script.Serialization;

namespace ConvertImage
{
    public partial class ParseSlides : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string SaveLocation = "";
            if ((ppt_file.PostedFile != null) && (ppt_file.PostedFile.ContentLength > 0))
            {
                string fn = System.IO.Path.GetFileName(ppt_file.PostedFile.FileName);
                SaveLocation = Server.MapPath("data") + "\\" + fn;
                try
                {
                    ppt_file.PostedFile.SaveAs(SaveLocation);
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

            PresentationEx pres = new PresentationEx(@SaveLocation);

            int count = pres.Slides.Count;

            string[] imageNameAry = new string[pres.Slides.Count];

            Bitmap bmp;
            string savePath = "";
            for (int i = 0; i < pres.Slides.Count; i++)
            {
                bmp = pres.Slides[i].GetThumbnail(1f, 1f);
                imageNameAry[i] = Guid.NewGuid().ToString();
                savePath = Server.MapPath("public/slides") + "\\" + imageNameAry[i] + ".jpg";
                bmp.Save(@savePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                bmp = pres.Slides[i].GetThumbnail(0.2f, 0.2f);
                savePath = Server.MapPath("public/slides") + "\\" + imageNameAry[i] + "_thumb.jpg";
                bmp.Save(@savePath, System.Drawing.Imaging.ImageFormat.Jpeg);
            }

            Response.Clear();
            var json = new JavaScriptSerializer().Serialize(imageNameAry);
            Response.Write(json);

            Response.End();
        }
    }
}