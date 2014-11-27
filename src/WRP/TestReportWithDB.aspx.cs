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
            string lsReportPath = @"TestReports/ReportWithOracle.rpt",
                lsServerName = "odbc_oracle", lsUserID = "ANA", lsPassword = "ANA", lsDatabaseName = "", lsPrefixDatabaseTable = "ANA.";
            lsReportPath = Server.MapPath(lsReportPath);
            ReportDocument lReportDocument;
            CrystalReport lGeneral_CR = new CrystalReport(lsServerName, lsUserID, lsPassword, lsDatabaseName, lsPrefixDatabaseTable);
            lReportDocument = lGeneral_CR.PrepareReport(lsReportPath);

            // appareance
            CrystalReportViewer1.HasCrystalLogo = false;
            CrystalReportViewer1.HasRefreshButton = true;
            CrystalReportViewer1.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None;

            // assign report documnt to CrystalReportViewer
            CrystalReportViewer1.ReportSource = lReportDocument;

        }

    }
}