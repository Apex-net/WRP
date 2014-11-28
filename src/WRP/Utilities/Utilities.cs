using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Data;
using System.Reflection;

namespace WebReportPreview.Utilities
{
    public class Common
    {
        /// <summary>
		/// Costruttore
		/// </summary>
        public Common()
		{
			//
			// TODO: Add constructor logic here
			//
		}

        /// <summary>
		/// Pulisce il nome del report da tutto ciò che non è il nome
		/// </summary>
		/// <param name="asReportPath"></param>
		/// <returns></returns>
		public string CleanReportPath(string asReportPath)
		{
			//
			string lsReportPath = asReportPath;
			int liLastBar, liLastBar1, liLastBar2 = 0;

			liLastBar1 = lsReportPath.LastIndexOf("\\");
			liLastBar2 = lsReportPath.LastIndexOf("/");

			if (liLastBar1 > 0 || liLastBar2 > 0)
			{
				liLastBar = Math.Max(liLastBar1, liLastBar2);
				lsReportPath = lsReportPath.Substring(liLastBar + 1);
			}
			return lsReportPath;

		}

        public string GetStatusBar()
        {
           
            string appo = System.Reflection.Assembly.GetExecutingAssembly().GetName().ToString();
            //Leggo le informazioni sul prodotto e sulla versione dall'AssemblyInfo.cs
            AssemblyProductAttribute productName = (AssemblyProductAttribute)AssemblyProductAttribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(AssemblyProductAttribute));
            string versionNumber = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            AssemblyCompanyAttribute companyDesc = (AssemblyCompanyAttribute)AssemblyCompanyAttribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(AssemblyCompanyAttribute));
            //Scrivo nella Status Bar le informazioni del WRP
            string statusbar = "";
            statusbar = companyDesc.Company + " " + productName.Product + " " + versionNumber;

            return statusbar;

        }

        /// <summary>
        /// Get Report Resource Name
        /// </summary>
        /// <param name="ReportPath"></param>
        /// <param name="ResourcePrefix"></param>
        /// <returns></returns>
        protected string ReportResourcesName(string ReportPath, string ResourcePrefix)
        {
            string reportResourcesName = "";
            reportResourcesName = CleanReportPath(ReportPath);
            if (ResourcePrefix != "" && ResourcePrefix != null)
            {
                reportResourcesName = ResourcePrefix + reportResourcesName;
            }
            return reportResourcesName;
        }


        /// <summary>
        /// Get Datatable Localizer From Resource
        /// </summary>
        /// <param name="ReportResource"></param>
        /// <returns></returns>
        public DataTable GetDataTableLocalizer(string ReportPath, string ResourcePrefix)
        {
            // DataTable
            DataTable tmpDt = new DataTable("myLoc");
            // Create two column and add to the DataTable.
            tmpDt.Columns.Add("Name", System.Type.GetType("System.String"));
            tmpDt.Columns.Add("Value", System.Type.GetType("System.String"));

            try
            {
                // 
                string reportResource = ReportResourcesName(ReportPath, ResourcePrefix);
                // Create Resource Manager
                System.Reflection.Assembly myAssembly = System.Reflection.Assembly.Load("App_GlobalResources");
                System.Resources.ResourceManager rm = new System.Resources.ResourceManager(String.Format("Resources.{0}", reportResource),
                                 myAssembly);
                // Get Resource Set
                //System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentUICulture;
                System.Resources.ResourceSet set = rm.GetResourceSet(ci, true, true);
                if (set != null)
                {
                    // Create datatable localization
                    System.Collections.IDictionaryEnumerator enumerator = set.GetEnumerator();
                    while (enumerator.MoveNext())
                    {
                        string key = (string)(enumerator.Key);
                        string value = (string)(enumerator.Value);
                        DataRow locDataRow;
                        locDataRow = tmpDt.NewRow();
                        locDataRow["Name"] = key;
                        locDataRow["Value"] = value;
                        tmpDt.Rows.Add(locDataRow);
                    }
                }
    	    }
            catch (Exception ex)
            {
                //
                string error = ex.Message;
		    }
            //
            return tmpDt;
        }

	}

}