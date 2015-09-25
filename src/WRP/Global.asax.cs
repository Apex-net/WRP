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
        }

        protected void Session_End(Object sender, EventArgs e)
		{
            //[!] todo: test (no in windows 10 machine)
            // 
            // Current verssion and Service Pack (SP 14) of SAP Crystal Reports, Developer Version for Visual Studio .NET do not support Windows 10 

            // Clean up report documents in sessione
            //
            try
            {
                // Test id sessione
                string id = Session.SessionID.ToString();
                // Cerco report documnent in session
                string prefixIdReportDocument = (string)Session["PrefixIdReportDocument"];
                string cleanUpRepDocEndSession = (string)Session["CleanUpRepDocEndSession"];
                if (cleanUpRepDocEndSession == "1" && prefixIdReportDocument != "")
                {
                    // se abiltato cleanup su sessione cerco i reportdocument validi e eseguo cleanup
                    for (int i = 0; i < Session.Contents.Count; i++)
                    {
                        string renamedItem;
                        renamedItem = Session.Keys[i];
                        int findPrefix = 0;
                        findPrefix = renamedItem.IndexOf(prefixIdReportDocument);
                        if (findPrefix == 0)
                        {
                            // Trovato
                            CrystalDecisions.CrystalReports.Engine.ReportDocument lReportDocument;
                            lReportDocument = (CrystalDecisions.CrystalReports.Engine.ReportDocument)Session[renamedItem];
                            if (lReportDocument != null)
                            {
                                // new:
                                WRP.Engine.CrystalReport.CloseReportDocument(lReportDocument);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                System.Diagnostics.Debug.WriteLine("Error DEBUG WRP - Session_End: " + error);
            }
        }
    }
}