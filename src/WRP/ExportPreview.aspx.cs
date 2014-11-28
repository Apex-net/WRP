using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using WRP.Engine;
using System.Threading;
using System.Globalization;

namespace WebReportPreview
{
    public partial class ExportPreview : System.Web.UI.Page
    {
        protected CrystalReport report;

        protected void Page_Load(object sender, System.EventArgs e)
        {
            // Put user code to initialize the page here

            if (!Page.IsPostBack)
            {
                ViewExportReport();
            }
            this.Title = "Export";

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
        /// View in page Exported Report.
        /// </summary>
        private void ViewExportReport()
        {
            // 
            string lsReportPath,
                lsSelectionFormula,
                lsServerName,
                lsUserDB,
                lsPasswordDB,
                lsDBName,
                lsTablePrefix,
                lsDisplayTreeview,
                lsDisplayToolbarCrystal,
                lsExportFormatType,
                lsDeleteExportFile,
                lsPhisicalReportPath;
            ReportDocument lReportDocument;
            ExportFormatType leftExportFormatType;
            bool lbOk;
            string lsContentType, lsUniqueIdentifier, lsReportFileExport, lsTempDir;
            DateTime ldt;
            DataTable ldtable = new DataTable();
            string lsCleanUpReportDocument;
            string lsCrEngineType, lsRasServerName, lsRasStreamingType;
            string lsExportedNameFile, lsExportAdvanced, lsExportAsAttachment;
            bool lbExportAsAttachment = false;
            DataTable ldtSrMod = new DataTable();
            string lsInitParamToEmptyString;
            string lsApplicationName;
            string lsLocalizeFromResource;
            string lsResourcePrefix;
            // Recupero i valori di parametrizzazione dalla session
            lsReportPath = (string)Session["ReportPath"];
            lsPhisicalReportPath = (string)Session["PhisicalReportPath"];
            lsSelectionFormula = (string)Session["sf"];
            if (Session["Parameters"] != null)
            {
                ldtable = (DataTable)Session["Parameters"];
            }
            else
            {
                ldtable = null;
            }
            lsServerName = (string)Session["ServerName"];
            lsUserDB = (string)Session["UserDB"];
            lsPasswordDB = (string)Session["PasswordDB"];
            lsDBName = (string)Session["DBName"];
            lsTablePrefix = (string)Session["TablePrefix"];
            lsDisplayTreeview = (string)Session["DisplayTreeview"];
            lsDisplayToolbarCrystal = (string)Session["DisplayToolbarCrystal"];
            lsExportFormatType = (string)Session["ExportFormatType"];
            lsDeleteExportFile = (string)Session["DeleteExportFile"];
            lsCleanUpReportDocument = (string)Session["CleanUpReportDocument"];
            lsCrEngineType = (string)Session["CrEngineType"];
            lsRasServerName = (string)Session["RasServerName"];
            lsRasStreamingType = (string)Session["RasStreamingType"];
            lsExportedNameFile = (string)Session["ExportedNameFile"];
            lsExportAdvanced = (string)Session["ExportAdvanced"];
            lsExportAsAttachment = (string)Session["ExportAsAttachment"];
            ldtSrMod = (DataTable)Session["SubReportModify"];
            lsInitParamToEmptyString = (string)Session["InitParamToEmptyString"];
            lsApplicationName = (string)Session["ApplicationName"];
            lsLocalizeFromResource = (string)Session["LocalizeFromResource"];
            lsResourcePrefix = (string)Session["ResourcePrefix"];
            if (lsExportAsAttachment == "1")
            {
                lbExportAsAttachment = true;
            }
            //
            if (lsReportPath != "" && lsReportPath != null)
            {
                // Recupero il path fisico del report
                if (lsPhisicalReportPath != "1")
                {
                    lsReportPath = Server.MapPath(lsReportPath);
                }
                //
                report = new CrystalReport(lsServerName, lsUserDB, lsPasswordDB, lsDBName, lsTablePrefix);
                // Inizializzo configurazioni di specializzazione sui sottoreport 
                if (ldtSrMod != null)
                {
                    report.SetSubReportModify(ldtSrMod);
                }
                // Inizializzo gli eventuali parametri a stringa vuota
                if (lsInitParamToEmptyString == "1")
                {
                    report.InitParametersToEmptyString = true;
                }

                if (Session["ForceDbPropertiesChange"] != null && (string)Session["ForceDbPropertiesChange"] == "1")
                    report.ForceDbPropertiesChange = true;

                // Elaboro il Report Document
                lReportDocument = report.PrepareReport(lsReportPath);
                lReportDocument = report.AddSelectionFormula(lReportDocument, lsSelectionFormula);
                lReportDocument = report.SetArrayParameterValue(lReportDocument, ldtable);
                // 
                if (lsLocalizeFromResource == "1")
                {
                    // Localize Report
                    Utilities.Common common = new Utilities.Common();
                    DataTable myDt = common.GetDataTableLocalizer(lsReportPath, lsResourcePrefix);
                    lReportDocument = report.ReportLocalizer(lReportDocument, myDt);
                }
                //
                // Export del Report Document
                //
                lsContentType = "";
                switch (lsExportFormatType)
                {
                    case "pdf":
                        //
                        leftExportFormatType = ExportFormatType.PortableDocFormat;
                        lsContentType = @"application/pdf";
                        break;
                    case "xls":
                        //
                        leftExportFormatType = ExportFormatType.Excel;
                        lsContentType = @"application/vnd.ms-excel";
                        break;
                    case "doc":
                        //
                        leftExportFormatType = ExportFormatType.RichText;
                        lsContentType = @"application/msword";
                        break;
                    case "rtf":
                        //
                        lsExportAdvanced = "1";
                        leftExportFormatType = ExportFormatType.RichText;
                        break;
                    case "rte":
                        //
                        lsExportAdvanced = "1";
                        leftExportFormatType = ExportFormatType.EditableRTF;
                        break;
                    case "xlr":
                        //
                        lsExportAdvanced = "1";
                        leftExportFormatType = ExportFormatType.ExcelRecord;
                        break;
                    case "txt":
                        //
                        lsExportAdvanced = "1";
                        lbExportAsAttachment = true; // download immediato
                        leftExportFormatType = ExportFormatType.Text;
                        break;
                    case "ttx":
                        //
                        lsExportAdvanced = "1";
                        lbExportAsAttachment = true; // download immediato
                        leftExportFormatType = ExportFormatType.TabSeperatedText;
                        break;
                    case "csv":
                        //
                        lsExportAdvanced = "1";
                        lbExportAsAttachment = true; // download immediato
                        leftExportFormatType = ExportFormatType.CharacterSeparatedValues;
                        break;
                    case "rpt":
                        //
                        lsExportAdvanced = "1";
                        lbExportAsAttachment = true; // Il txt viene gestito solo nel caso di download immediato
                        leftExportFormatType = ExportFormatType.CrystalReport;
                        break;
                    default:
                        //
                        leftExportFormatType = ExportFormatType.PortableDocFormat;
                        lsContentType = @"application/pdf";
                        break;
                }
                // Export
                if (lsExportAdvanced == "1")
                {
                    // Utilizzo metodo dedicato di Report Document 
                    //
                    lReportDocument.ExportToHttpResponse(leftExportFormatType, Response, lbExportAsAttachment, lsExportedNameFile);
                }
                else
                {
                    // Export e Response "manuale" 
                    lsTempDir = Server.MapPath("Temp");
                    //
                    lsUniqueIdentifier = lsExportedNameFile;
                    lsUniqueIdentifier += Session.SessionID.ToString();
                    ldt = DateTime.Now;
                    lsUniqueIdentifier += ldt.ToString("yyyyMMddHHmmss");
                    //
                    lsReportFileExport = "";
                    lbOk = report.ExportDocument(lReportDocument, leftExportFormatType, lsTempDir, lsUniqueIdentifier, ref lsReportFileExport);
                    //
                    string lsAttachment = "";
                    if (lbExportAsAttachment)
                    {
                        lsAttachment = "attachment";
                    }
                    //
                    if (lbOk == true)
                    {
                        // Output al browser con utilizzo di Trasmitfile()
                        Response.Clear();
                        Response.ContentType = lsContentType;
                        Response.AppendHeader("Content-Disposition", lsAttachment);
                        Response.TransmitFile(lsReportFileExport);
                        Response.End();
                        // Cancello Export file
                        if (lsDeleteExportFile == "1")
                        {
                            System.IO.File.Delete(lsReportFileExport);
                        }
                    }
                }
                // Rilascio risorse report document
                // 
                CrystalReport.CloseReportDocument(lReportDocument);

            }
        }
    }
}