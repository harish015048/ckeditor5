using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Web;

namespace CMCai.Actions
{
    public class WordGenerator
    {
        public string ExporttoWord(string HtmlString)
        {
            string filename = string.Empty;           
            filename = "Summary_Report_" + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");                     
            //HttpContext.Current.Response.Clear();
            //HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + filename + ".xls");
            //HttpContext.Current.Response.ContentType = "application/vnd.xls";
            //HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache); // not necessarily required
            //HttpContext.Current.Response.Charset = "";
            //HttpContext.Current.Response.Output.Write(HtmlString);
            using (StreamWriter file = new StreamWriter(System.Web.Hosting.HostingEnvironment.MapPath(@"~\ReportDocs\" + filename + ".doc"), true, Encoding.UTF8))
            {

                file.WriteLine(HtmlString);
            }

            string user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            //string path = System.Web.Hosting.HostingEnvironment.MapPath("~/ReportDocs/");
            //string filePath = string.Empty;
            //Deny writing to the file
            AddDirectorySecurity(System.Web.Hosting.HostingEnvironment.MapPath(@"~\ReportDocs\" + filename + ".doc"), user, FileSystemRights.Write, AccessControlType.Deny);

            return filename + ".doc";
        }

        public static void AddDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            // Create a new DirectoryInfo object.
            DirectoryInfo dInfo = new DirectoryInfo(FileName);


            // Get a DirectorySecurity object that represents the 
            // current security settings.
            DirectorySecurity dSecurity = dInfo.GetAccessControl();


            // Add the FileSystemAccessRule to the security settings. 
            dSecurity.AddAccessRule(new FileSystemAccessRule(Account,
            Rights, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None,
            ControlType));


            // Set the new access settings.
            dInfo.SetAccessControl(dSecurity);


        }

        public static void SetFileReadAccess(string FileName, bool SetReadOnly)
        {
            FileInfo fInfo = new FileInfo(FileName);

            // Set the IsReadOnly property.
            fInfo.IsReadOnly = SetReadOnly;

        }


        // Returns wether a file is read-only.
        public static bool IsFileReadOnly(string FileName)
        {
            // Create a new FileInfo object.
            FileInfo fInfo = new FileInfo(FileName);

            // Return the IsReadOnly property value.
            return fInfo.IsReadOnly;

        }

        public string ExporttoWord(string HtmlString, string repTyp)
        {
            string filename = string.Empty;

            filename = repTyp + DateTime.Now.ToString("MMddyyyy_hhmmssmmmtt");


            using (StreamWriter file = new StreamWriter(System.Web.Hosting.HostingEnvironment.MapPath(@"~\PublishFiles\" + filename + ".doc"), true, Encoding.UTF8))
            {
                file.WriteLine(HtmlString);
            }

            return filename + ".doc";
        }
    }
}