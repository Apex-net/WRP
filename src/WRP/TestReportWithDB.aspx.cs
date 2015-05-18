using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WRP.Engine;
using CrystalDecisions.CrystalReports.Engine;

namespace WebReportPreview
{
    public partial class TestReportWithDB : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            // set a information message
            string is64bitProcess = System.Environment.Is64BitProcess.ToString();
            Label1.Text = "Below, there is a DB connected report " + " - " + "Is it a 64bit Process? " + is64bitProcess;
            Label1.Font.Bold = true;

        }

        protected void Page_Init(object sender, EventArgs e)
        {

            // set up report
            string lsReportPath = @"TestReports/ReportWithOracleBocconi.rpt",
                lsServerName = "odbc_oracle_bocconi", lsUserID = "UNIBOCCONI_AGE20_DEV", lsPassword = "age", lsDatabaseName = "", 
                lsPrefixDatabaseTable = "UNIBOCCONI_AGE20_DEV.",
                lsSelectionFormula = "UpperCase ({AGE_UTENTI.Nome}) = 'ANDREA'";
            lsReportPath = Server.MapPath(lsReportPath);
            ReportDocument lReportDocument;
            CrystalReport lGeneral_CR = new CrystalReport(lsServerName, lsUserID, lsPassword, lsDatabaseName, lsPrefixDatabaseTable);
            lReportDocument = lGeneral_CR.PrepareReport(lsReportPath);
            lReportDocument = lGeneral_CR.AddSelectionFormula(lReportDocument, lsSelectionFormula);

            // appareance
            CrystalReportViewer1.HasCrystalLogo = false;
            CrystalReportViewer1.HasRefreshButton = true;
            CrystalReportViewer1.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None;

            // assign report documnt to CrystalReportViewer
            CrystalReportViewer1.ReportSource = lReportDocument;

        }

    }
}