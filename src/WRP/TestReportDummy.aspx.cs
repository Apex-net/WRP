using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebReportPreview
{
    public partial class TestReportDummy : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            // label
            string is64bitProcess = System.Environment.Is64BitProcess.ToString();
            Label1.Text = "Below, there is a Crystal Report" + " - " + "Is it a 64bit Process? " + is64bitProcess;
            Label1.Font.Bold = true;

            // Report viewer
            CrystalReportViewer1.HasCrystalLogo = false;
            CrystalReportViewer1.HasRefreshButton = true;
            CrystalReportViewer1.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None;

        }
    }
}