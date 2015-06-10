using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using WRP.API.Models;
using WRP.Engine;

namespace WRP.API.Controllers
{
    public class ExportController : ApiController
    {

        [HttpGet]
        public IHttpActionResult GetHelpForReportConfig()
        {
            var result = new ReportConfig
            {
                ReportPath = "E' il percorso assoluto del report da visualizzare, completo del nome del file",

                ConnectionServerName = "Indica il nome del DNS Odbc o dell'alias DB che verrà utilizzato per connettersi al DB",
                ConnectionUserDB = "Nome dell'utente che verrà utilizzato per connettersi al DB",
                ConnectionPasswordDB = "Password utilizzata per connettersi al DB",
                ConnectionDBName = "Eventuale db alternativo a quello associato di default del login",
                ConnectionTablePrefix = "Eventuale prefisso al nome delle table nel report",

                SelectionFormula = "Selection Formula, contiene la formula che viene passata al file di report",

                ExportPath = "Percorso della directory di output",
                ExportNamefile = "Nome file output",
                ExportFormat = "Formato di export - pdf, xls, doc, rtf (rich text), rte (editable rich text), xlr (ExcelRecords), txt, ttx (TabSeparatedText), csv (CharacterSeparatedValues), rpt (Crystal Reports)"

            };
            return this.Ok(result);
        }


        [HttpPut]
        public IHttpActionResult ExportReport([FromBody]ReportConfig reportConfig)
        {
            string exportReturn = "";
            exportReturn = ExportDocumentHelper(reportConfig);
            
            var result = new ResultExport
            {
                ResultCode = exportReturn == "Success" ? 0 : 1,
                ResultMessage = exportReturn == "Success" ? "Success" : exportReturn
            };

            return this.Ok(result);
        }

        public string ExportDocumentHelper(ReportConfig reportConfig)
        {
            string exportPath = "";
            string errorMessage = "";

            try
            {
                CrystalReport report = new CrystalReport(serverName: reportConfig.ConnectionServerName,
                userID: reportConfig.ConnectionUserDB,
                password: reportConfig.ConnectionPasswordDB,
                databaseName: reportConfig.ConnectionDBName,
                prefixDatabaseTable: reportConfig.ConnectionTablePrefix);

                exportPath = report.ExportDocumentSimple(PathReport: reportConfig.ReportPath,
                SelectionFormula: reportConfig.SelectionFormula,
                PathExport: reportConfig.ExportPath,
                Namefile: reportConfig.ExportNamefile,
                ExportFormat: reportConfig.ExportFormat);
            }
            catch (Exception ex)
            {
                errorMessage = String.Format("WRP.API.Controllers.ExportReport: {0}", ex.Message);
            }

            return string.IsNullOrEmpty(errorMessage) ? "Success" : errorMessage;
        }
       
    }
}
