using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace CMCai.Actions
{
    public class ExcelGenerator
    {
        public string ExporttoExcel(string HtmlString, short repTyp)
        {
            string filename = string.Empty;
            if (repTyp == 1)
            {
                filename = "Registration_Report_" + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");
            }
            if (repTyp == 2)
            {
                filename = "Commitment_Report_" + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");
            }
            if (repTyp == 3)
            {
                filename = "AuditTrailReport_" + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");
            }
            if (repTyp == 5)
            {
                filename = "eIFUFeedbackReport_" + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");
            }
            if (repTyp == 6)
            {
                filename = "LibraryReport_" + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");
            }
            if (repTyp == 7)
            {
                filename = "CompanyList_" + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");
            }

            //HttpContext.Current.Response.Clear();
            //HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + filename + ".xls");
            //HttpContext.Current.Response.ContentType = "application/vnd.xls";
            //HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache); // not necessarily required
            //HttpContext.Current.Response.Charset = "";
            //HttpContext.Current.Response.Output.Write(HtmlString);
            using (StreamWriter file = new StreamWriter(System.Web.Hosting.HostingEnvironment.MapPath(@"~\Uploads\" + filename + ".xls"), true, Encoding.UTF8))
            {
                file.WriteLine(HtmlString);
            }

            return filename + ".xls";
        }
    }
}