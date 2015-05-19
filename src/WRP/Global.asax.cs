using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.SessionState;

namespace WebReportPreview
{
    public class Global : System.Web.HttpApplication
    {
        protected void Application_Start(object sender, EventArgs e)
        {
        }
        protected void Application_Error(Object sender, EventArgs e)
        {

            // cattura l'ultimo errore verificatosi nal server
			Exception objErr = Server.GetLastError().GetBaseException();
            System.Diagnostics.Debug.WriteLine("Error DEBUG WRP: " + objErr.Message);
            // [!] todo 
            // transfer control to error page
            /*
            if ((string)Session["HandleGlobalError"] == "1")
                Server.Transfer("Errors.aspx");
             */


        }
    }
}