using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Data;

namespace WRP.Engine
{
    /// <summary>
    /// 
    /// </summary>
    class CrystalReport
    {
        // Class Variables
		protected string csServerName, csUserID, csPassword, csDatabaseName;
		protected string csPrefixDatabaseTable;
		protected DataTable csSubReportModify;
        protected bool initParametersToEmptyString;
        protected bool csForceDbPropertiesChange;

        public bool ForceDbPropertiesChange
        {
            get
            {
                return csForceDbPropertiesChange;
            }
            set
            {
                csForceDbPropertiesChange = value;
            }
        }

        /// <summary>
        /// Inizializzo a empty string tutti i valori dei parametri del report document
        /// </summary>
        public bool InitParametersToEmptyString
        {
            set
			{
                initParametersToEmptyString = value;

			}
			get
			{
                return initParametersToEmptyString;
			}
        }

		/// <summary>
		/// Costruttore
		/// </summary>
		public CrystalReport()
		{
			//
			// TODO: Add constructor logic here
			//
			SetConnectionInfo ("", "", "", "", "");
			//
			SetDefaults();
		}

		/// <summary>
		///  Costruttore: overloading del costruttore standard
		///  Setta alcune condizioni di Connection info per il report e alcuni default
		/// </summary>
		/// <param name="serverName">Nome del server</param>
		/// <param name="userID">Username di accesso al database</param>
		/// <param name="password">Parssword di accesso al database</param>
		/// <param name="databaseName">Nome del database</param>
		/// <param name="prefixDatabaseTable">Prefisso standard dei nomi delle tabelle</param>
		/// <example><code>
		/// CrystalReport lGeneral_CR = new CrystalReport(lsServerName, lsUserID, lsPassword, lsDatabaseName, lsPrefixDatabaseTable);
		/// </code></example>
		public CrystalReport(string serverName, string userID, string password, string databaseName, string prefixDatabaseTable)
		{
			//
			// TODO: Add constructor logic here
			//
			SetConnectionInfo (serverName, userID, password, databaseName, prefixDatabaseTable);
			//
			SetDefaults();
		}

		/// <summary>
		/// Il metodo, chiamato dal costruttore, setta alcune condizioni di Connection info per i reports gestiti dalla classe
		/// </summary>
		/// <param name="serverName">Nome del server</param>
		/// <param name="userID">Username di accesso al database</param>
		/// <param name="password">Parssword di accesso al database</param>
		/// <param name="databaseName">Nome del database</param>
		/// <param name="prefixDatabaseTable">Prefisso standard dei nomi delle tabelle</param>
		public void SetConnectionInfo (string serverName, string userID, string password, string databaseName, string prefixDatabaseTable)
		{
			//
			csServerName = serverName; 
			csUserID = userID; 
			csPassword = password; 
			csDatabaseName = databaseName;
			csPrefixDatabaseTable = prefixDatabaseTable;
		}
		/// <summary>
		/// Permette di settare la configurazione e il comportamento di sottoreports a run-time
		/// </summary>
		/// <param name="subReportModify">datatable contenente le modifiche di comportamento sul sottoreport</param>
		public void SetSubReportModify (DataTable subReportModify)
		{
			//
			if (subReportModify != null)
			{
				if (subReportModify.Rows.Count == 0)
				{
					subReportModify = null;
				}
			}
			csSubReportModify = subReportModify; 
		}

		/// <summary>
		/// Il metodo setta alcuni defaults
		/// </summary>
		protected void SetDefaults ()
		{
            initParametersToEmptyString = false;
            csForceDbPropertiesChange = false;
        }

		/// <summary>
		/// Restituisce un oggetto "report document", dato il path completo di un report di CR, 
		/// applicando le info di connection per ogni tabella e sottoreport incluso.
		/// Può connettersi anche ad un server RAS (Unmanaged o Managed)
		/// </summary>
		/// <param name="reportPath">Percorso del report</param>
		/// <returns></returns>
		/// <example><code>
		/// ReportDocument lReportDocument;
		/// CrystalReport lGeneral_CR = new CrystalReport(lsServerName, lsUserID, lsPassword, lsDatabaseName, lsPrefixDatabaseTable);
		/// lReportDocument = lGeneral_CR.PrepareReport(lsReportPath);
		/// </code></example>
		public ReportDocument PrepareReport(string reportPath) 
		{ 
			ReportDocument crReportDocument;
			crReportDocument = new ReportDocument();
			//
			try
			{
			    crReportDocument.Load(reportPath, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy);
				SetConnectionDB(crReportDocument);
			}
			catch (Exception ex)
			{
				throw (new Exception("CrystalReport.PrepareReport: " + ex.Message, ex));
			}
			return crReportDocument; 
		}

		/// <summary>
		/// Applica le info di connection DB ad un report document
		/// </summary>
		/// <param name="ardReportDocument"></param>
		protected void SetConnectionDB(ReportDocument ardReportDocument)
		{
			ReportDocument crReportDocument;
			ConnectionInfo crConnectionInfo; 
			Database crDatabase; 
			Tables crTables; 
			
			string suffixDatabaseTable;
			//
			crReportDocument =  ardReportDocument;
			//
			try
			{

                crDatabase = crReportDocument.Database;
                crTables = crDatabase.Tables;
                crConnectionInfo = new ConnectionInfo();
                crConnectionInfo.AllowCustomConnection = true;
                crConnectionInfo.ServerName = csServerName;
                if (csDatabaseName != "")
                {
                    crConnectionInfo.DatabaseName = csDatabaseName;
                }
                crConnectionInfo.UserID = csUserID;
                crConnectionInfo.Password = csPassword;

                if (csForceDbPropertiesChange)
                {
                    ApplyLoginInfo(crReportDocument);
                }
                else
                {
                    foreach (CrystalDecisions.CrystalReports.Engine.Table aTable in crReportDocument.Database.Tables)
                    {
                        TableLogOnInfo crTableLogOnInfo = aTable.LogOnInfo;
                        crTableLogOnInfo.ConnectionInfo = crConnectionInfo;
                        aTable.ApplyLogOnInfo(crTableLogOnInfo);

                        if (csPrefixDatabaseTable != "")
                        {
                            suffixDatabaseTable = aTable.Location.Substring(aTable.Location.LastIndexOf(".") + 1);
                            aTable.Location = csPrefixDatabaseTable + suffixDatabaseTable;
                        }
                    }
                }

                // Setta Connection per eventuali sottoreports
				SetConnectionForSubReport (crReportDocument); 			
			}
			catch (Exception ex)
			{
				throw (new Exception("CrystalReport.SetConnectionDB: " + ex.Message, ex));
			}
		}

        /// <summary>
        /// Forza il cambio delle proprietà di connessione al db su tutte sul report principale, su tutte le tabelle e su tutti i sottoreport
        /// </summary>
        /// <param name="document"></param>
        private void ApplyLoginInfo(ReportDocument document)
        {
            TableLogOnInfo info = null;
            try
            {
                //Credentials
                info = new TableLogOnInfo();
                info.ConnectionInfo.AllowCustomConnection = true;
                info.ConnectionInfo.ServerName = csServerName;
                info.ConnectionInfo.DatabaseName = csDatabaseName;
                info.ConnectionInfo.Password = csPassword;
                info.ConnectionInfo.UserID = csUserID;

                // Apply to connections, tables and sub-reports
                document.SetDatabaseLogon(info.ConnectionInfo.UserID, info.ConnectionInfo.Password, info.ConnectionInfo.ServerName, info.ConnectionInfo.DatabaseName, false);

                // Other connections?
                foreach (CrystalDecisions.Shared.IConnectionInfo connection in document.DataSourceConnections)
                {
                    connection.SetConnection(csServerName, csDatabaseName, false);
                    connection.SetLogon(csUserID, csPassword);
                    connection.LogonProperties.Set("Data Source", csServerName);
                    connection.LogonProperties.Set("Initial Catalog", csDatabaseName);
                }

                // Only do this to the main report (can't do it to sub reports)                
                if (!document.IsSubreport)
                {
                    // Apply to subreports
                    foreach (ReportDocument rd in document.Subreports)
                    {
                        ApplyLoginInfo(rd);
                    }
                }

                // Apply to tables
                foreach (CrystalDecisions.CrystalReports.Engine.Table table in document.Database.Tables)
                {
                    TableLogOnInfo tableLogOnInfo = table.LogOnInfo;

                    tableLogOnInfo.ConnectionInfo = info.ConnectionInfo;
                    table.ApplyLogOnInfo(tableLogOnInfo);
                    //if (!table.TestConnectivity())
                        //Debug.WriteLine("Failed to apply log in info for Crystal Report");
                }

                /*
                try
                {
                    // Break it all down
                    document.VerifyDatabase();
                }
                catch (LogOnException excLogon)
                {
                    //Debug.WriteLine(excLogon.Message);
                }*/
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Failed to apply login information to the report - " +
                    ex.Message);
            }
        }


		/// <summary>
		/// Applica le condizioni di restrizione a un oggetto "report document"
		/// </summary>
		/// <param name="ardReportDocument">oggetto "report document"</param>
		/// <param name="selectionFormula">Formula di selezione</param>
		/// <returns>Restituisce un oggetto "report document", con applicate le condizioni di restizione</returns>
		public ReportDocument AddSelectionFormula(ReportDocument ardReportDocument, string selectionFormula) 
		{ 
			ReportDocument crReportDocument;
			crReportDocument = ardReportDocument; 
			if (selectionFormula != "" && selectionFormula != null)
			{
				crReportDocument.DataDefinition.RecordSelectionFormula = selectionFormula;
			}
			return crReportDocument; 
		}

		/// <summary>
		/// Setta ad un report document tutti i valori dei parameters con i valori della Datatable
		/// </summary>
		/// <param name="ardReportDocument">Report document in input</param>
		/// <param name="parameterDataTable">Datatable, nella la prima dimensione sono memorizzati i nomi dei parameter, nella seconda i valori</param>
		/// <returns>Restituisce un oggetto "report document", con settato tutti i valori dei parameters</returns>
		public ReportDocument SetArrayParameterValue(ReportDocument ardReportDocument, DataTable parameterDataTable)
		{ 
			ReportDocument crReportDocument;
			crReportDocument = ardReportDocument;
			DataTable tmpParam;
			int count = 0;
            string currentName, currentValue, currentValueInit;
            if (initParametersToEmptyString == true)
            {
                currentValueInit = "";
                foreach (CrystalDecisions.Shared.ParameterField pf in crReportDocument.ParameterFields)
                {
                    if (pf.ParameterValueType == ParameterValueKind.StringParameter)
                    {
                        ParameterDiscreteValue pdv = new ParameterDiscreteValue();
                        pdv.Value = currentValueInit;
                        crReportDocument.SetParameterValue(pf.Name, pdv);
                    }
                }
            }
			if (parameterDataTable != null)
			{
				tmpParam = parameterDataTable;			
				count = tmpParam.Rows.Count; 
				for( int li = 0; li < count; li++ )
				{
					currentName = tmpParam.Rows[li][0].ToString();
					currentValue = tmpParam.Rows[li][1].ToString();
					if (currentName != "" && currentName != null)
					{
						crReportDocument.SetParameterValue(currentName, currentValue);
					}
					else
					{
						break;
					}
				}
			}
			return crReportDocument; 
		}

        /// <summary>
		/// Applica le info di connection DB ad un eventuale sottoreport di un report document
		/// </summary>
		/// <param name="aReportDocument"></param>
		protected void SetConnectionForSubReport (ReportDocument aReportDocument)
		{ 
			Sections crSections; 
			ReportDocument crReportDocument, crSubreportDocument;
			SubreportObject crSubreportObject;
			ReportObjects crReportObjects; 
			ConnectionInfo crConnectionInfo; 
			Database crDatabase; 
			Tables crTables; 
			TableLogOnInfo crTableLogOnInfo;
			DataTable tmpSubReportModify;
			string tempPrefixDatabaseTable;
			string tempSubreportName;
			string curSrName, curServerName, curUserID, curPassword, curDatabaseName, curPrefixDatabaseTable;
			int count;
		
			crReportDocument = aReportDocument;
			tmpSubReportModify = csSubReportModify;
			tempPrefixDatabaseTable = csPrefixDatabaseTable;

			// set the sections object to the current report's section 
			crSections = crReportDocument.ReportDefinition.Sections; 
			// loop through all the sections to find all the report objects 
			foreach (Section crSection in crSections) 
			{ 
				crReportObjects = crSection.ReportObjects; 
				//loop through all the report objects in there to find all subreports 
				foreach (ReportObject crReportObject in crReportObjects) 
				{
					if (crReportObject.Kind == ReportObjectKind.SubreportObject) 
					{
						try
						{
							crSubreportObject = (SubreportObject) crReportObject; 
							// open the subreport object and logon as for the general report or as subreport modified
							tempSubreportName = crSubreportObject.SubreportName;
							crSubreportDocument = crSubreportObject.OpenSubreport(tempSubreportName); 
							crDatabase = crSubreportDocument.Database; 
							crTables = crDatabase.Tables; 
							// loop sulle varie tables per settare le connection info
							foreach (CrystalDecisions.CrystalReports.Engine.Table aTable in crTables) 
							{
								//
								crConnectionInfo = new ConnectionInfo();
								//
								if (tmpSubReportModify != null)
								{
									// check per eventuale sottoreport "to modify"
									count = 0;
									count = tmpSubReportModify.Rows.Count; 
									for( int li = 0; li < count; li++ )
									{
										curSrName = tmpSubReportModify.Rows[li][0].ToString();
										curServerName = tmpSubReportModify.Rows[li][1].ToString();
										curUserID = tmpSubReportModify.Rows[li][2].ToString();
										curPassword = tmpSubReportModify.Rows[li][3].ToString();
										curDatabaseName = tmpSubReportModify.Rows[li][4].ToString();
										curPrefixDatabaseTable = tmpSubReportModify.Rows[li][5].ToString();
										if (tempSubreportName == curSrName)
										{
											crConnectionInfo.ServerName = curServerName; 
											crConnectionInfo.UserID = curUserID; 
											crConnectionInfo.Password = curPassword;
											if (curDatabaseName != "")
											{
												crConnectionInfo.DatabaseName = curDatabaseName;
											}
											tempPrefixDatabaseTable = curPrefixDatabaseTable;
											break;
										}
										else 
										{
											crConnectionInfo.ServerName = csServerName; 
											crConnectionInfo.UserID = csUserID; 
											crConnectionInfo.Password = csPassword;
											if (csDatabaseName != "")
											{
												crConnectionInfo.DatabaseName = csDatabaseName;
											}
											tempPrefixDatabaseTable = csPrefixDatabaseTable;
										}
									}
								}
								else
								{
									crConnectionInfo.ServerName = csServerName; 
									crConnectionInfo.UserID = csUserID; 
									crConnectionInfo.Password = csPassword;
									if (csDatabaseName != "")
									{
										crConnectionInfo.DatabaseName = csDatabaseName;
									}
									tempPrefixDatabaseTable = csPrefixDatabaseTable;
								}
								// Setto qui le connection info per la table
								crTableLogOnInfo = aTable.LogOnInfo;
								crTableLogOnInfo.ConnectionInfo = crConnectionInfo;
								aTable.ApplyLogOnInfo(crTableLogOnInfo); 
								if (tempPrefixDatabaseTable != "")
								{
									aTable.Location = tempPrefixDatabaseTable + aTable.Location.Substring(aTable.Location.LastIndexOf(".") + 1);
								}
							}
						}
						catch (Exception ex)
						{
							throw (new Exception("CrystalReport.SetConnectionForSubReport: " + ex.Message, ex));
						}
					}
				}
			}
		}

		/// <summary>
		/// Esporta un oggetto "report document", in un determinato formato. 
		/// </summary>
		/// <param name="ardReportDocument">Oggetto "report document"</param>
		/// <param name="expFormatType">Formato di export</param>
		/// <param name="tempDir">Percorso della directory temporanea</param>
		/// <param name="identifier"></param>
		/// <param name="reportFileExport"></param>
		/// <returns>Ritorna True in caso di export completato con successo</returns>
		public bool ExportDocument (ReportDocument ardReportDocument, ExportFormatType expFormatType, 
			string tempDir, string identifier,	ref string reportFileExport)
		{ 
			string lReportFileExport, extensionFile;
			ExportOptions expOptions;
			DiskFileDestinationOptions diskFileDestOptions;
			ReportDocument lReportDocument;
			ExportFormatType lExpFormatType;
			bool isOK;

			lReportFileExport = "";
			extensionFile = "";
			lExpFormatType = expFormatType;
			isOK = false;
			lReportDocument = ardReportDocument;
			if	(	(expFormatType == ExportFormatType.PortableDocFormat) 
				||	(expFormatType == ExportFormatType.Excel)
				||	(expFormatType == ExportFormatType.ExcelRecord)
				||	(expFormatType == ExportFormatType.RichText)
				||	(expFormatType == ExportFormatType.Text)
				||	(expFormatType == ExportFormatType.CrystalReport)
				||	(expFormatType == ExportFormatType.HTML32)
				||	(expFormatType == ExportFormatType.HTML40)
				||	(expFormatType == ExportFormatType.WordForWindows)
				||	(expFormatType == ExportFormatType.EditableRTF)
				||	(expFormatType == ExportFormatType.TabSeperatedText)
				||	(expFormatType == ExportFormatType.CharacterSeparatedValues)
				)
			{
				// Export Format Options and Extension file
				switch(expFormatType) 
				{
					case ExportFormatType.PortableDocFormat:
						//
						extensionFile = ".pdf";
						break;
					case ExportFormatType.Excel:
						//
						extensionFile = ".xls";
						break;
					case ExportFormatType.ExcelRecord:
						//
						extensionFile = ".xls";
						break;
					case ExportFormatType.RichText:
						//
						extensionFile = ".rtf";
						break;
					case ExportFormatType.Text:
						//
						extensionFile = ".txt";
						break;
					case ExportFormatType.CrystalReport:
						//
						extensionFile = ".rpt";
						break;
					case ExportFormatType.HTML32:
					case ExportFormatType.HTML40:
						//
						extensionFile = ".html";
						// Nel caso dell'Html occore definire le HTMLFormatOptions come ExportOptions del lReportDocument
						CrystalDecisions.Shared.HTMLFormatOptions htmlOption = new CrystalDecisions.Shared.HTMLFormatOptions();
						htmlOption.HTMLBaseFolderName = tempDir;
						htmlOption.HTMLFileName = identifier + extensionFile;
						//Aggiunge una barra di navigazione alle pagine
						htmlOption.HTMLHasPageNavigator = true;
						//Abilita la separazione delle pagine
						htmlOption.HTMLEnableSeparatedPages = true;
						lReportDocument.ExportOptions.FormatOptions = htmlOption;
						break;
					case ExportFormatType.WordForWindows:
						//
						extensionFile = ".doc";
						break;
					case ExportFormatType.EditableRTF:
						//
						extensionFile = ".rtf";
						break;
					case ExportFormatType.TabSeperatedText:
						//
						extensionFile = ".txt";
						break;
					case ExportFormatType.CharacterSeparatedValues:
						//
						extensionFile = ".csv";
						//
						CharacterSeparatedValuesFormatOptions lCsvOp = new CharacterSeparatedValuesFormatOptions();
						lCsvOp.PreserveDateFormatting = true;
						lCsvOp.PreserveNumberFormatting = true;
						lCsvOp.Delimiter = "";
						lCsvOp.SeparatorText = ";";
						lReportDocument.ExportOptions.FormatOptions = lCsvOp;
						break;
					default:
						//
						extensionFile = ".pdf";
						break;
				}
				// Export options
				lReportFileExport =  tempDir + @"\\" + identifier + extensionFile;
				diskFileDestOptions = new DiskFileDestinationOptions();
				diskFileDestOptions.DiskFileName = lReportFileExport;
				expOptions = lReportDocument.ExportOptions;
				expOptions.DestinationOptions = diskFileDestOptions;
				expOptions.ExportDestinationType = ExportDestinationType.DiskFile;
				expOptions.ExportFormatType = lExpFormatType;
				// Export
				//
				try
				{
					lReportDocument.Export();
					isOK = true;
				}
				catch (Exception ex)
				{
					throw (new Exception("CrystalReport.ExportDocument: " + ex.Message, ex));
				}
			}
			reportFileExport = lReportFileExport;

			return isOK;
        }

        /// <summary>
        /// Esportazione, semplificata, di un report in un formato supportato.
        /// </summary>
        /// <param name="PathReport">Path fisico report</param>
        /// <param name="SelectionFormula">Formula di selezione</param>
        /// <param name="PathExport">Percorso della directory di output</param>
        /// <param name="Namefile">Nome file output</param>
        /// <param name="ExportFormat">Formato di export
        /// Tipi supportati: 
        /// pdf, xls, doc, rtf (rich text), rte (editable rich text), xlr (ExcelRecords), txt, ttx (TabSeparatedText), csv (CharacterSeparatedValues), rpt (Crystal Reports)
        /// </param>
        /// <returns>Restituisce il path fisico completo del documento esportato</returns>
        /// <example><code>
		/// CrystalReport lGeneral_CR = new CrystalReport(_ServerName, _UserID, _Password, _DatabaseName, _PrefixDatabaseTable);
        /// string retPathExport = lGeneral_CR.ExportDocumentSimple(_ReportPath, _SelectionFormula, _PathExport, _Namefile, _ExportFormat);
		/// </code></example>
        public string ExportDocumentSimple(string PathReport, string SelectionFormula, string PathExport, string Namefile, string ExportFormat)
        {
            string retPathExport = "";
            ReportDocument reportDoc;
            ExportFormatType expFormatType;

            // Creazione report document locale
            reportDoc = PrepareReport(PathReport);
            reportDoc = AddSelectionFormula(reportDoc, SelectionFormula);

            // export differenziato in base al tipo di formato prescelto 
            switch (ExportFormat)
            {
                case "pdf":
                    //
                    expFormatType = ExportFormatType.PortableDocFormat;
                    break;
                case "xls":
                    //
                    expFormatType = ExportFormatType.Excel;
                    break;
                case "doc":
                    //
                    expFormatType = ExportFormatType.WordForWindows;
                    break;
                case "rtf":
                    //
                    expFormatType = ExportFormatType.RichText;
                    break;
                case "rte":
                    //
                    expFormatType = ExportFormatType.EditableRTF;
                    break;
                case "xlr":
                    //
                    expFormatType = ExportFormatType.ExcelRecord;
                    break;
                case "txt":
                    //
                    expFormatType = ExportFormatType.Text;
                    break;
                case "ttx":
                    //
                    expFormatType = ExportFormatType.TabSeperatedText;
                    break;
                case "csv":
                    //
                    expFormatType = ExportFormatType.CharacterSeparatedValues;
                    break;
                case "rpt":
                    //
                    expFormatType = ExportFormatType.CrystalReport;
                    break;
                default:
                    //
                    expFormatType = ExportFormatType.PortableDocFormat;
                    break;
            }
            // Export 
            //
            string uniqueIdentifier = Namefile;
            ExportDocument(reportDoc, expFormatType, PathExport, uniqueIdentifier, ref retPathExport);
            //

            return retPathExport;
        }

        /// <summary>
        /// Localizza un report document in base ad un data table contente le stringhe da localizzare
        /// </summary>
        /// <par\am name="aReportDocument"></param>
        /// <param name="aDt"></param>
        /// <returns></returns>
        public ReportDocument ReportLocalizer(ReportDocument aReportDocument, DataTable aDt)
        {
            ReportDocument lReportDocument = aReportDocument;
            DataTable tmpDt = aDt;

            try
            {
                if (tmpDt != null)
                {
                    int count = tmpDt.Rows.Count;
                    for (int li = 0; li < count; li++)
                    {
                        string key = tmpDt.Rows[li][0].ToString();
                        string value = tmpDt.Rows[li][1].ToString();
                        lReportDocument = AssignKeysToReport(lReportDocument, key, value);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(String.Format("CrystalReport.ReportLocalizer: {0}", ex.Message), ex);
            }

            return lReportDocument;
        }

        /// <summary>
        /// Cambia il valore di un Textobject o di un Formulafield nel Reportdocument (translate)
        /// </summary>
        /// <param name="myReportDocument"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private ReportDocument AssignKeysToReport(ReportDocument myReportDocument, string key, string value)
        {
            ReportDocument lReportDocument = myReportDocument;

            string reportObjectName = key;
            string reportObjectValue = value;

            // Set value of text object in report
            try
            {
                ReportObject reportObject = lReportDocument.ReportDefinition.ReportObjects[reportObjectName];
                if (reportObject != null)
                {
                    TextObject myLabel = (TextObject)reportObject;
                    myLabel.Text = reportObjectValue;
                }
            }
            catch (Exception ex)
            {
                // Translation not found (No error handling)
                string error = ex.Message;
            }

            // Set value of formula field in report
            try
            {
                FormulaFieldDefinition reportObject = lReportDocument.DataDefinition.FormulaFields[reportObjectName];
                if (reportObject != null)
                {
                    FormulaFieldDefinition myLabel = (FormulaFieldDefinition)reportObject;
                    myLabel.Text = reportObjectValue;
                }
            }
            catch (Exception ex)
            {
                // Translation not found (No error handling)
                string error = ex.Message;
            }

            // Search on  Subreports
            //
            Sections crSections;
            ReportDocument crReportDocument, crSubreportDocument;
            SubreportObject crSubreportObject;
            ReportObjects crReportObjects;
            string tempSubreportName;

            string subReportObjectName = key;
            string subRreportObjectValue = value;

            crReportDocument = lReportDocument;
            crSections = crReportDocument.ReportDefinition.Sections;

            // loop through all the sections to find all the report objects 
            foreach (Section crSection in crSections)
            {
                crReportObjects = crSection.ReportObjects;
                //loop through all the report objects in there to find all subreports 
                foreach (ReportObject crReportObject in crReportObjects)
                {
                    if (crReportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        crSubreportObject = (SubreportObject)crReportObject;
                        // open the subreport object and logon as for the general report or as subreport modified
                        tempSubreportName = crSubreportObject.SubreportName;
                        crSubreportDocument = crSubreportObject.OpenSubreport(tempSubreportName);

                        // Set value of text object in report
                        try
                        {
                            ReportObject subreportObject = crSubreportDocument.ReportDefinition.ReportObjects[subReportObjectName];
                            if (subreportObject != null)
                            {
                                TextObject mySubLabel = (TextObject)subreportObject;
                                mySubLabel.Text = subRreportObjectValue;
                            }
                        }
                        catch (Exception ex)
                        {
                            // No error handling (no translate found)
                            string error = ex.Message;
                        }

                        // Set value of formula filed in report
                        try
                        {
                            FormulaFieldDefinition subreportObject = crSubreportDocument.DataDefinition.FormulaFields[subReportObjectName];
                            if (subreportObject != null)
                            {
                                FormulaFieldDefinition mySubLabel = (FormulaFieldDefinition)subreportObject;
                                mySubLabel.Text = subRreportObjectValue;
                            }
                        }
                        catch (Exception ex)
                        {
                            // Translation not found (No error handling)
                            string error = ex.Message;
                        }
                    }
                }
            }

            return lReportDocument;
        }

        
        /// <summary>
        /// Stampa diretta
        /// </summary>
        /// <param name="aReportDocument">Report Document</param>
        /// <param name="aPrinterName">Nome stampante, se vuoto viene utilizzata la stampante di default</param>
        public static void DirectPrint(ReportDocument aReportDocument, string aPrinterName)
        {
            ReportDocument myReportDocument = aReportDocument;
            string myPrinterName = aPrinterName; 

            try
            {
                //
                if (myPrinterName != "")
                {
                    myReportDocument.PrintOptions.PrinterName = myPrinterName;
                }
                //
                myReportDocument.PrintToPrinter(1, false, 0, 0);

            }
            catch (Exception ex)
            {
                throw (new Exception(String.Format("CrystalReport.DirectPrint: {0}", ex.Message), ex));
            }

        }

        /// <summary>
        /// Rilascia risorse ReportDocument (close, dispose, ...)
        /// </summary>
        /// <param name="aReportDocument">Report Document</param>
        public static void CloseReportDocument(ReportDocument aReportDocument)
        {
            ReportDocument myReportDocument = aReportDocument;

            try
            {
                if (myReportDocument != null)
                {
                    if (myReportDocument.IsLoaded == true)
                    {
                        myReportDocument.Close();
                        myReportDocument.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                throw (new Exception(String.Format("CrystalReport.CloseReportDocument: {0}", ex.Message), ex));
            }

        }
	}
}
