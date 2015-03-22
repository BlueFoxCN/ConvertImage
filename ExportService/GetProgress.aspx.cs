using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ConvertImage
{
    public partial class GetProgress : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                Console.WriteLine(Request.Params);
                string job_id = Request.Params["job_id"];
                if (job_id == null)
                    return;
                double progress = Progress.progress[job_id];
                if (progress == 1)
                {
                    Progress.progress.Remove(job_id);
                }
                Response.Clear();
                Response.Write(progress);
            }
            catch (KeyNotFoundException)
            {
                Response.Clear();
                Response.Write("0");
            }
            catch (Exception es)
            {
                Response.Clear();
                Response.Write("error:" + es);
            }
            Response.End();
        }
    }
}