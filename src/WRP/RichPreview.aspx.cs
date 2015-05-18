using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.Web;
using WRP.Engine;
using System.Threading;
using System.Globalization;
using System.Configuration;

namespace WebReportPreview
{
    public partial class RichPreview : System.Web.UI.Page
    {

        protected CrystalReport ireport;
        protected ReportDocument iReportDocument;
        protected string iUniqueIdentifier;

        /// <summary>
        /// Page Init
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Init(object sender, EventArgs e)
        {

            // 
            if (!this.IsPostBack)
            {
                // Configuro aspetto visuale viewer
                ConfigReportViewer();
                // Inizializzo report document
                InizializeReport();
            }
            else
            {
                string lsCleanUpReportDocument = (string)Session["CleanUpReportDocument"];
                if (lsCleanUpReportDocument == "1")
                {
                    // Inizializzo report document
                    InizializeReport();
                }
                else
                {
                    // Recupero report document da session
                    string reportId = Page.Request.Form.Get("__ApexNetReportId");
                    iUniqueIdentifier = reportId;
                    iReportDocument = (ReportDocument)Session[reportId];
                }
            }

            // Assegno report document a viewer
            AssignReportSource();

            // Imposto il titolo della pagina
            SetTitlePage();

            // Gestione campi nascosti di persistenza stato alternativi al viewstate
            MyRegisterHiddenFields();

        }

        /// <summary>
        /// Set the culture and UI culture for the ASP.NET Web page programmatically
        /// </summary>
        protected override void InitializeCulture()
        {

            string lsOverrideCulture, lsOverridedUICulture, lsOverridedCulture;

            lsOverrideCulture = (string)Session["OverrideCulture"];
            lsOverridedUICulture = (string)Session["OverridedUICulture"];
            lsOverridedCulture = (string)Session["OverridedCulture"];

            // Set the culture and UI culture
            if (lsOverrideCulture == "1")
            {
                if (lsOverridedCulture != "")
                {
                    Thread.CurrentThread.CurrentCulture =
                          CultureInfo.CreateSpecificCulture(lsOverridedCulture);
                }
                if (lsOverridedUICulture != "")
                {
                    Thread.CurrentThread.CurrentUICulture = new
                            CultureInfo(lsOverridedUICulture);
                }
            }
            base.InitializeCulture();
        }

        /// <summary>
        /// Page Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// Clean up
        /// </summary>
        protected void Page_UnLoad(object sender, System.EventArgs e)
        {
            // CleanUp Report Document
            string lsCleanUpReportDocument = (string)Session["CleanUpReportDocument"];
            if (iReportDocument != null)
            {
                if (lsCleanUpReportDocument == "1")
                {
                    CrystalReport.CloseReportDocument(iReportDocument);
                }

            }
        }

        /// <summary>
        /// Initialize Report Document
        /// </summary>
        protected void InizializeReport()
        {
            // Inizialize report
            string lsReportPath,
                lsSelectionFormula,
                lsServerName,
                lsUserDB,
                lsPasswordDB,
                lsDBName,
                lsTablePrefix,
                lsPhisicalReportPath,
                lsInitParamToEmptyString,
                lsLocalizeFromResource,
                lsResourcePrefix;
            ReportDocument lReportDocument;
            DataTable ldtable, ldtSrMod;
            int lNumParameters;
            ldtable = new DataTable();
            ldtSrMod = new DataTable();

            //Recupero i valori dalla session
            lsReportPath = (string)Session["ReportPath"];
            lsPhisicalReportPath = (string)Session["PhisicalReportPath"];
            lsSelectionFormula = (string)Session["sf"];
            if (Session["Parameters"] != null)
            {
                ldtable = (DataTable)Session["Parameters"];
                lNumParameters = ldtable.Rows.Count;
            }
            else
            {
                lNumParameters = 0;
                ldtable = null;
            }
            //
            lsServerName = (string)Session["ServerName"];
            lsUserDB = (string)Session["UserDB"];
            lsPasswordDB = (string)Session["PasswordDB"];
            lsDBName = (string)Session["DBName"];
            lsTablePrefix = (string)Session["TablePrefix"];
            //
            lsInitParamToEmptyString = (string)Session["InitParamToEmptyString"];
            ldtSrMod = (DataTable)Session["SubReportModify"];
            //
            lsLocalizeFromResource = (string)Session["LocalizeFromResource"];
            lsResourcePrefix = (string)Session["ResourcePrefix"];
            //
            if (lsReportPath != "" && lsReportPath != null)
            {
                // Recupero il path fisico del report
                if (lsPhisicalReportPath != "1")
                {
                    lsReportPath = Server.MapPath(lsReportPath);
                }
                // Istanzio classe di lavoro
                ireport = new CrystalReport(lsServerName, lsUserDB, lsPasswordDB, lsDBName, lsTablePrefix);
                // Inizializzo configurazioni di specializzazione sui sottoreport 
                if (ldtSrMod != null)
                {
                    ireport.SetSubReportModify(ldtSrMod);
                }
                // Inizializzo gli eventuali parametri a stringa vuota
                if (lsInitParamToEmptyString == "1")
                {
                    ireport.InitParametersToEmptyString = true;
                }
                // Lavoro con ReportDocument        

                if (Session["ForceDbPropertiesChange"] != null && (string)Session["ForceDbPropertiesChange"] == "1")
                    ireport.ForceDbPropertiesChange = true;

                lReportDocument = ireport.PrepareReport(lsReportPath);
                lReportDocument = ireport.AddSelectionFormula(lReportDocument, lsSelectionFormula);
                lReportDocument = ireport.SetArrayParameterValue(lReportDocument, ldtable);
                if (lsLocalizeFromResource == "1")
                {
                    // Localize Report
                    Utilities.Common common = new Utilities.Common();
                    DataTable myDt = common.GetDataTableLocalizer(lsReportPath, lsResourcePrefix);
                    lReportDocument = ireport.ReportLocalizer(lReportDocument, myDt);
                }
                // 
                // Salvo report document come istanza
                iReportDocument = lReportDocument;

                // Salvo in session
                // Gestisco un report id per poter avere in sessione più 
                // report document e permettere l'apertura contemporanea di
                // più report sullo stesso session ID. 
                // Memorizzo il valore in hidden field per poterlo ritrovare al postback
                string lsPrefixIdReportDocument = (string)Session["PrefixIdReportDocument"];
                string uniqueIdentifier = "";
                uniqueIdentifier = lsPrefixIdReportDocument;
                DateTime ldt = DateTime.Now;
                uniqueIdentifier += ldt.ToString("yyyyMMddHHmmss");
                iUniqueIdentifier = uniqueIdentifier;
                Session[uniqueIdentifier] = lReportDocument;

                // Gestione dinamica del titolo
                if ((string)Session["EnableTitle"] == "1" && (string)Session["Title"] != "" && (string)Session["Title"] != null)
                {
                    ViewState["ReportTitle"] = (string)Session["Title"];
                }
                else
                {
                    ViewState["ReportTitle"] = ConfigurationManager.AppSettings["Title"].ToString();
                }

            }
        }

        /// <summary>
        /// Assign report document to viewer source
        /// </summary>
        protected void AssignReportSource()
        {
            //
            if (iReportDocument != null)
            {
                richCrystalReportViewer.ReportSource = iReportDocument;
            }
        }


        /// <summary>
        /// Configure visual aspects
        /// </summary>
        protected void ConfigReportViewer()
        {
            string lsDisplayToolbarCrystal,
                lsHasCrystalLogo,
                lsPrintMode,
                lsHasToggleGroupTreeButton,
                lsHasExportButton,
                lsHasRefreshButton,
                lsHasPrintButton,
                lsHasViewList,
                lsHasDrillUpButton,
                lsHasPageNavigationButtons,
                lsHasGotoPageButton,
                lsHasSearchButton,
                lsHasZoomFactorList,
                lsHasToggleParameterPanelButton,
                lsToolPanelView;
            //
            lsDisplayToolbarCrystal = (string)Session["DisplayToolbarCrystal"];
            lsHasCrystalLogo = (string)Session["HasCrystalLogo"];
            lsPrintMode = (string)Session["PrintMode"];
            lsHasToggleGroupTreeButton = (string)Session["HasToggleGroupTreeButton"];
            lsHasExportButton = (string)Session["HasExportButton"];
            lsHasRefreshButton = (string)Session["HasRefreshButton"];
            lsHasPrintButton = (string)Session["HasPrintButton"];
            lsHasViewList = (string)Session["HasViewList"];
            lsHasDrillUpButton = (string)Session["HasDrillUpButton"];
            lsHasPageNavigationButtons = (string)Session["HasPageNavigationButtons"];
            lsHasGotoPageButton = (string)Session["HasGotoPageButton"];
            lsHasSearchButton = (string)Session["HasSearchButton"];
            lsHasZoomFactorList = (string)Session["HasZoomFactorList"];
            lsHasToggleParameterPanelButton = (string)Session["HasToggleParameterPanelButton"];
            lsToolPanelView = (string)Session["ToolPanelView"];
            //
            // Configurazione Report Viewer
            // 
            if (lsDisplayToolbarCrystal == "0")
            {
                richCrystalReportViewer.DisplayToolbar = false;
            }
            if (lsPrintMode == "X")
            {
                richCrystalReportViewer.PrintMode = PrintMode.ActiveX;
            }
            else
            {
                richCrystalReportViewer.PrintMode = PrintMode.Pdf;
            }
            if (lsHasToggleGroupTreeButton == "0")
            {
                richCrystalReportViewer.HasToggleGroupTreeButton = false;
            }
            if (lsHasExportButton == "0")
            {
                richCrystalReportViewer.HasExportButton = false;
            }
            if (lsHasRefreshButton == "0")
            {
                richCrystalReportViewer.HasRefreshButton = false;
            }
            if (lsHasPrintButton == "0")
            {
                richCrystalReportViewer.HasPrintButton = false;
            }
            if (lsHasDrillUpButton == "0")
            {
                richCrystalReportViewer.HasDrillUpButton = false;
            }
            if (lsHasPageNavigationButtons == "0")
            {
                richCrystalReportViewer.HasPageNavigationButtons = false;
            }
            if (lsHasGotoPageButton == "0")
            {
                richCrystalReportViewer.HasGotoPageButton = false;
            }
            if (lsHasSearchButton == "0")
            {
                richCrystalReportViewer.HasSearchButton = false;
            }
            if (lsHasZoomFactorList == "0")
            {
                richCrystalReportViewer.HasZoomFactorList = false;
            }
            if (lsHasCrystalLogo == "1")
            {
                richCrystalReportViewer.HasCrystalLogo = true;
            }
            if (lsHasToggleParameterPanelButton == "0")
            {
                richCrystalReportViewer.HasToggleParameterPanelButton = false;
            }
            switch (lsToolPanelView)
            {
                case "P":
                    richCrystalReportViewer.ToolPanelView = ToolPanelViewType.ParameterPanel;
                    break;
                case "G":
                    richCrystalReportViewer.ToolPanelView = ToolPanelViewType.GroupTree;
                    break;
                default:
                    richCrystalReportViewer.ToolPanelView = ToolPanelViewType.None;
                    break;
            }
        }

        /// <summary>
        /// Set browser page title
        /// </summary>
        protected void SetTitlePage()
        {
            // Imposto il titolo del report sul nome pagina
            this.Title = (string)ViewState["ReportTitle"];
        }

        /// <summary>
        /// Manage my hidden fields
        /// </summary>
        protected void MyRegisterHiddenFields()
        {
            // Identifier for my reportdocument
            this.ClientScript.RegisterHiddenField("__ApexNetReportId", iUniqueIdentifier);
        }

    }
}