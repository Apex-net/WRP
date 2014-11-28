using System;
using System.Collections.Generic;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

namespace WebReportPreview
{
    /// <summary>
    /// Entry point page 
    /// </summary>
    public partial class ApexNetPreview : System.Web.UI.Page
    {
        string applicationName;
        string reportPath;

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {
                // Carico ReportPath
                reportPath = (string)Request.QueryString["ReportPath"];

                // Se ReportPath non è vuoto e non è nullo allora comincio altrimenti non faccio nulla
                if (reportPath != "" && reportPath != null)
                {
                    // Carico ApplicationName
                    applicationName = (string)Request.QueryString["ApplicationName"];
                    // Se ApplicationName non l'ho trovato nella QueryString carico quello di default dal Web.config
                    if (applicationName == "" || applicationName == null)
                    {
                        applicationName = ConfigurationManager.AppSettings["SectionDBdefault"].ToString();
                    }
                    // Carico nella Session le chiavi di "appSettings" del Web.config
                    LoadFromWebConfig("appSettings");
                    // Carico nella Session le chiavi della sezione "ApplicationName" del Web.config
                    LoadFromWebConfig(applicationName);
                    // Carico nella Session i parametri presenti nella QueryString
                    LoadFromQueryString();
                    // Carico nella Session gli eventuali valori da Post Http 
                    LoadFromHttpPost();
                    // Carica i parametri p(n) dentro un'unica variabile Session["Parameters"]
                    GetParameters();
                    // Carica le modifiche di connessione per i subreports sr(n)
                    GetSubReportModify();
                    // Log
                    SaveQueryStringAndPostInLog();

                    // Lancio pagina preview del tipo indicato  
                    string lsPreviewType = (string)Session["PreviewType"];
                    bool lbPreserveFormTransfer = false;
                    switch (lsPreviewType)
                    {
                        case "RIC":
                            //
                            Response.Redirect("RichPreview.aspx", false);
                            //
                            break;
                        case "EXP":
                            //
                            Response.Redirect("ExportPreview.aspx", false);
                            break;
                        default:
                            //
                            Server.Transfer("ExportPreview.aspx", lbPreserveFormTransfer);
                            break;
                    }
                }
                else
                {
                    Response.Write("Non è stato specificato alcun report!");
                }
            }

        }

        /// <summary>
        /// Carica dentro la Session tutte le chiavi presenti in un nodo del web.config
        /// </summary>
        /// <param name="nodeName"></param>
        public void LoadFromWebConfig(string nodeName)
        {
            string chiave, valore;
            var nodeToLoad = ConfigurationManager.GetSection(nodeName) as NameValueCollection;
            for (int i = 0; i < nodeToLoad.Count; i++)
            {
                chiave = nodeToLoad.GetKey(i).ToString();
                valore = nodeToLoad.Get(chiave);
                Session[chiave] = valore;
            }
        }

        /// <summary>
        /// Carica dentro la Session tutte le variabili specificate nella QueryString
        /// </summary>
        public void LoadFromQueryString()
        {
            string chiave, valore;
            NameValueCollection queryStringCollection = Request.QueryString;

            for (int i = 0; i < queryStringCollection.Count; i++)
            {
                chiave = queryStringCollection.GetKey(i).ToString();
                valore = queryStringCollection.Get(chiave);
                Session[chiave] = valore;
            }

            // Trattamento differenziato per la chiave "sf"
            string lsKeySpecific, lsValSpecific;
            lsKeySpecific = "sf";
            lsValSpecific = queryStringCollection.Get(lsKeySpecific);
            if (lsValSpecific == "" || lsValSpecific == null)
            {
                Session[lsKeySpecific] = "";
            }
        }

        /// <summary>
        /// Carica dentro la Session tutte le variabili specificate in eventuale HttpPost
        /// </summary>
        public void LoadFromHttpPost()
        {
            string getValuesFromHttpPost = "";
            getValuesFromHttpPost = (string)Session["GetValuesFromHttpPost"];
            if (getValuesFromHttpPost == "1")
            {
                string chiave, valore;
                NameValueCollection requestFormCollection = Request.Form;
                // Leggo dati in post e scrivo session
                for (int i = 0; i < requestFormCollection.Count; i++)
                {
                    chiave = requestFormCollection.GetKey(i).ToString();
                    valore = requestFormCollection.Get(chiave);
                    Session[chiave] = valore;
                }
            }
        }

        /// <summary>
        /// Carica le modifiche di connessione
        /// per i subreports sr(n) dentro un'unica variabile Session["SubReportModify"]
        /// </summary>
        protected void GetSubReportModify()
        {
            // Carica le modifiche per i subreports sr(n) dentro un'unica variabile Session["SubReportModify"]
            DataTable ldtSrMod = new DataTable();
            DataRow ldrowSrMod;
            int lmi = 1;
            // Create five columns and add to the DataTable.
            ldtSrMod.Columns.Add("SrName", System.Type.GetType("System.String"));
            ldtSrMod.Columns.Add("ServerName", System.Type.GetType("System.String"));
            ldtSrMod.Columns.Add("UserID", System.Type.GetType("System.String"));
            ldtSrMod.Columns.Add("Password", System.Type.GetType("System.String"));
            ldtSrMod.Columns.Add("DatabaseName", System.Type.GetType("System.String"));
            ldtSrMod.Columns.Add("PrefixDatabaseTable", System.Type.GetType("System.String"));
            // 
            string lsN = (string)Session["sN" + lmi.ToString()];
            string lsA = (string)Session["sA" + lmi.ToString()];
            // 
            while (lsN != null && lsA != null)
            {
                // Create DataRow object, and add to DataTable.
                ldrowSrMod = ldtSrMod.NewRow();
                // 
                string lServerName, lUserID, lPassword, lDatabaseName, lPrefixDatabaseTable;
                var nodeToLoad = ConfigurationManager.GetSection(lsA) as NameValueCollection;
                lServerName = nodeToLoad.Get("ServerName");
                lUserID = nodeToLoad.Get("UserDB");
                lPassword = nodeToLoad.Get("PasswordDB");
                lDatabaseName = nodeToLoad.Get("DBName");
                lPrefixDatabaseTable = nodeToLoad.Get("TablePrefix");
                //
                ldrowSrMod["SrName"] = lsN;
                ldrowSrMod["ServerName"] = lServerName;
                ldrowSrMod["UserID"] = lUserID;
                ldrowSrMod["Password"] = lPassword;
                ldrowSrMod["DatabaseName"] = lDatabaseName;
                ldrowSrMod["PrefixDatabaseTable"] = lPrefixDatabaseTable;
                //
                ldtSrMod.Rows.Add(ldrowSrMod);
                //
                Session.Remove("sN" + lmi.ToString());
                Session.Remove("sA" + lmi.ToString());
                //
                lmi = lmi + 1;
                lsN = (string)Session["sN" + lmi.ToString()];
                lsA = (string)Session["sA" + lmi.ToString()];
            }
            Session["SubReportModify"] = ldtSrMod;
        }

        /// <summary>
        /// Carica i parametri p(n) dentro un'unica variabile Session["Parameters"]
        /// </summary>
        protected void GetParameters()
        {
            DataTable ldtab = new DataTable();
            DataRow ldrow;
            int li = 1;
            ldtab.Columns.Add("Name", System.Type.GetType("System.String"));
            ldtab.Columns.Add("Value", System.Type.GetType("System.String"));
            string lp = (string)Session["p" + li.ToString()];
            string lv = (string)Session["v" + li.ToString()];
            while (lp != null && lv != null)
            {
                ldrow = ldtab.NewRow();
                ldrow["Name"] = lp;
                ldrow["Value"] = lv;
                ldtab.Rows.Add(ldrow);
                Session.Remove("p" + li.ToString());
                Session.Remove("v" + li.ToString());

                li = li + 1;
                lp = (string)Session["p" + li.ToString()];
                lv = (string)Session["v" + li.ToString()];
            }
            Session["Parameters"] = ldtab;
        }

        /// <summary>
        /// Scrivo keys query string in log se gestito
        /// </summary>
        protected void SaveQueryStringAndPostInLog()
        {
            string writeIntoLog = "";

            writeIntoLog = (string)Session["WriteIntoLog"];

            if (writeIntoLog == "1")
            {
                // Leggi query string
                string url = Request.Url.ToString();
                string keyValueLog = "Url: " + url + Environment.NewLine;
                // Keys
                keyValueLog += "Keys Query String: " + Environment.NewLine;
                string chiave, valore;
                NameValueCollection queryStringCollection = Request.QueryString;
                for (int i = 0; i < queryStringCollection.Count; i++)
                {
                    chiave = queryStringCollection.GetKey(i).ToString();
                    valore = queryStringCollection.Get(chiave);
                    keyValueLog += chiave + ": " + valore + Environment.NewLine;
                }
                // Leggi eventuale post
                string getValuesFromHttpPost = "";
                getValuesFromHttpPost = (string)Session["GetValuesFromHttpPost"];
                if (getValuesFromHttpPost == "1")
                {
                    keyValueLog += "Keys PostHttp: " + Environment.NewLine;
                    string chiavePost, valorePost;
                    NameValueCollection requestFormCollection = Request.Form;
                    // Leggo dati in post e scrivo session
                    for (int i = 0; i < requestFormCollection.Count; i++)
                    {
                        chiavePost = requestFormCollection.GetKey(i).ToString();
                        valorePost = requestFormCollection.Get(chiavePost);
                        keyValueLog += chiavePost + ": " + valorePost + Environment.NewLine;
                    }
                }
                // Scrivi in log
                // [!]
                /*
                string lsMsg = keyValueLog;
                string lsEventLog = System.Configuration.ConfigurationSettings.AppSettings["EventLog"];
                string lsEventLogSource = System.Configuration.ConfigurationSettings.AppSettings["EventLogSource"];
                EvLog lev = new EvLog(lsEventLog, lsEventLogSource);
                lev.WriteEvLog(lsMsg, System.Diagnostics.EventLogEntryType.Information);
                 */
            }

        }

    }
}