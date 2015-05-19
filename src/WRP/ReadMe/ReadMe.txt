************************* Note per il richiamo di ApexNetPreview.aspx (entry point di Web Report Preview)

***** Note generali:

E' possibile "comandare" l'applicativo attraverso la pagina ApexNetPreview.aspx (entry point dell'applicativo).
Tale possibilità viene fornita attraverso la query string dell'url della pagina stessa e/o eventuali valori passati con post http.

La pagina ApexNetPreview "carica" le variabili presenti nella query string e del post direttamente nella session.
I parametri non specificati nella modalità standard vengono letti dal Web.config, alcuni dal nodo-sezione
specificato da "AplicationName", altri dal nodo "appSettings", secondo lo schema della tabella sottostante.
L'ordine in cui vengono caricati i parametri è: "appSettings", sezione, query string, post http; ciò significa che i
parametri contenuti in "appSettings" rappresentano il default, i parametri della sezione specializzano il
default per ogni applicativo, attraverso la query string e/o il post è possibile personalizzare l'apertura di ogni
singolo report.

Es: 
http://localhost/WRP/ApexNetPreview.aspx?ReportPath=TestReports/ReportWithOracleBocconi.rpt&PreviewType=EXP&ExportFormatType=pdf&sf=UpperCase ({AGE_UTENTI.Nome}) = 'ANDREA'
http://localhost/WRP/ApexNetPreview.aspx?ReportPath=TestReports/ReportWithOracleBocconi.rpt&PreviewType=RIC&sf=UpperCase ({AGE_UTENTI.Nome}) = 'ANDREA'
http://localhost/WRP/ApexNetPreview.aspx?ReportPath=TestReports/ReportSqlserverCRM.rpt&ApplicationName=CRM&PreviewType=RIC

Per tutti i parametri previsti il meccamismo è del tipo chiave e valore.

I parametri previsti sono:
---------------------------------------------------------------------------------------------
|                               |               | Post    | Query   |      Web.config       |
|    Nome Parametro             |     Tipo      | Http    | String  | section | appSettings |
|-------------------------------------------------------------------------------------------|
| ApplicationName               | String        |   NO    |   SI    |   NO    |     NO      |
| ReportPath                    | String        |   NO    |   SI    |   NO    |     NO      |
| PhisicalReportPath            | String        |   SI    |   SI    |   SI    |     SI      |
| PreviewType                   | {EXP,RIC}     |   SI    |   SI    |   SI    |     SI      |
| ServerName                    | String        |   SI    |   SI    |   SI    |     SI      |
| UserDB                        | String        |   SI    |   SI    |   SI    |     SI      |
| PasswordDB                    | String        |   SI    |   SI    |   SI    |     SI      |
| DBName                        | String        |   SI    |   SI    |   SI    |     SI      |
| TablePrefix                   | String        |   SI    |   SI    |   SI    |     SI      |
| sf                            | String        |   SI    |   SI    |   NO    |     NO      |
| ExportFormatType              | {vedi dett.)  |   SI    |   SI    |   SI    |     SI      |
| DeleteExportFile              | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| CleanUpReportDocument         | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| p1 .. pn                      | String        |   SI    |   SI    |   NO    |     NO      |
| v1 .. vn                      | String        |   SI    |   SI    |   NO    |     NO      |
| DisplayToolbarCrystal         | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| PrintMode                     | {X,P}         |   SI    |   SI    |   SI    |     SI      |
| HasToggleGroupTreeButton      | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasExportButton               | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasRefreshButton              | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasPrintButton                | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasViewList                   | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasDrillUpButton              | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasPageNavigationButtons      | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasGotoPageButton             | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasSearchButton               | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasZoomFactorList             | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HasCrystalLogo                | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| ExportedNameFile              | String        |   SI    |   SI    |   SI    |     SI      |
| ExportAdvanced                | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| ExportAsAttachment            | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| HandleGlobalError             | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| ViewErrors                    | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| WriteIntoLog                  | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| SupportMail                   | String        |   NO    |   NO    |   SI    |     SI      |
| EnableTitle                   | {0,1}	        |   NO    |   NO    |   SI    |     SI      |
| Title                         | String        |   SI    |   SI    |   SI    |     SI      |
| sN1 .. sNn                    | String        |   SI    |   SI    |   NO    |     NO      |
| sA1 .. sAn                    | String        |   SI    |   SI    |   NO    |     NO      |
| InitParamToEmptyString        | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| CleanUpRepDocEndSession       | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| PrefixIdReportDocument        | String        |   SI    |   SI    |   SI    |     SI      |
| HasToggleParameterPanelButton | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| ToolPanelView                 | {P,G,""}      |   SI    |   SI    |   SI    |     SI      |
| LocalizeFromResource          | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| ResourcePrefix                | String        |   SI    |   SI    |   SI    |     SI      |
| GetValuesFromHttpPost         | {0,1}         |   NO    |   SI    |   SI    |     SI      |
| OverrideCulture               | {0,1}         |   SI    |   SI    |   SI    |     SI      |
| OverridedUICulture            | String        |   SI    |   SI    |   SI    |     SI      |
| OverridedCulture              | String        |   SI    |   SI    |   SI    |     SI      |
| ForceDbPropertiesChange       | String        |   NO    |   NO    |   SI    |     SI      |
---------------------------------------------------------------------------------------------

I parametri obbligatori sono il ReportPath e ApplicationName (devono essere in query string), 
il resto è opzionale.

*** Analisi dettagliata dei vari parametri:

  ApplicationName:          Indica a ApexNetPreview.aspx quale sezione del Web.config caricare. Se non viene
                            passato nella Query String, la pagina caricherà la sezione indicata dal valore
                            della chiave "SectionDBdefault" del nodo "appSettings" del Web.config.
  
  ReportPath:               E' il percorso (relativo o assoluto) del report da visualizzare, completo del nome
                            del file.

  PhisicalReportPath:       Indica se il percorso indicato in "ReportPath" è fisico (1) o virtuale (0)
  
  PreviewType:              Indica il modo in cui verrà visualizzato il report
                              - EXP: Modalità ExportPreview,  il report viene esportato in un dato formato
                                (vedi parametro "ExportFormatFile") ed aperto con l'applicazione predefinita
                                per quel formato. 
                              - RIC: Modalità RichPreview, il report viene visualizzato all'interno del browser,
                                navigabile per pagina con funzioni di drill.
  
  ServerName:               Indica il nome del DNS Odbc o dell'alias DB che verrà utilizzato per connettersi al DB.
  
  UserDB:                   Nome dell'utente che verrà utilizzato per connettersi al DB.
  
  PasswordDB:               Password utilizzata per connettersi al DB.
  
  DBName:                   Eventuale db alternativo a quello associato di default del login
  
  TablePrefix:              Eventuale prefisso al nome delle table nel report
  
  sf:                       Selection Formula, contiene la formula che viene passata al file di report.
  
  ExportFormatType:         Nel caso di ExportPreview, indica in che formato sarà esportato il report.
                            Valori previsti: 
                            pdf, xls, doc, rtf (rich text), rte (editable rich text),
                            xlr (ExcelRecords), txt, ttx (TabSeparatedText), csv (CharacterSeparatedValues), 
                            rpt (Crystal Reports)
  
  DeleteExportFile:         Flag che indica se il file esportato sarà cancellato (1) o meno (0) alla chiusura
                            della pagina.
  
  CleanUpReportDocument:    Flag che indica se il REPORT DOCUMENT del report venga ricostruito (1) o recuperato dalla session (0)
                            ad ogni round trip al server; 
  
  p1 .. pn:                 Specifica il nome del parametro n-esimo che verrà passato al report.
  
  v1 .. vn:                 Specifica il valore del parametro n-esimo che verrà passato al report.

  DisplayToolbarCrystal:    Nel caso di RichPreview, indica se verrà visualizzata (1) o meno (0) la CrystalToolbar.

  PrintMode:                Nel caso di RichPreview, indica se la modalità di stampa è PDF (P) o direttamente
                            a stampante tramite ActiveX (X).
  
  HasToggleGroupTreeButton Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  HasExportButton          Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  HasRefreshButton         Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)  

  HasPrintButton           Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)  

  HasDrillUpButton         Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  HasPageNavigationButtons Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  HasGotoPageButton        Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  HasSearchButton          Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  HasZoomFactorList        Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  HasCrystalLogo           Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  ExportedNameFile         Nome del file di export

  ExportAdvanced           Indica se verrà utilizzato (1) o meno (0) la modalità di export avanzata

  ExportAsAttachment       Indica se il download del file partirà immediatamente (1) o in una nuova finestra browser (0)
  
  HandleGlobalError        Indica se deve (1) o non deve (0) essere visualizzata la pagina di gestione degli errori personalizzata
  
  ViewErrors               Indica se la pagina personalizzata di visualizzazione degli errori deve (1) o non deve (0) 
                           visualizzare la descrizione dell'errore catturato
  
  WriteIntoLog             Indica se l'applicazione deve (1) o non deve (0) scrivere gli errori nell'event Log 
                           Le configurazioni per la scrittura possono essere personalizzate con le chiavi EventLog e EventLogSource
                           La chiave EventLogSource deve essere presente come chiave nel registro.
  
  SupportMail              E-mail del supporto tecnico dell'applicativo

  EnableTitle              Indica se l'applicazione deve (1) o non deve (0) supportare i titoli dinamici dei report

  Title                    Titolo di default della pagina visualizzato nel browser

  sN1 .. sNn:              Specifica il nome del sottoreport n-esimo di cui si vogliono modificare le proprietà di connessione DB
  
  sA1 .. sAn:              Specifica il nome dell'Application name (section) utilizzato dal sottoreport n-esimo di cui sopra
  
  InitParamToEmptyString   Inizializzo gli eventuali parametri a stringa vuota

  CleanUpRepDocEndSession  Indica se deve (1) o non deve (0) esssere fatto il clean-up di tutti i reports gestiti allo scadere della sessione applicativa

  PrefixIdReportDocument   Prefisso identificativo Report Document per clean up di End Session (utilizzata solo se CleanUpRepDocEndSession = 1)

  HasToggleParameterPanelButton Indica se il corrispondente item, su toolbar di RichPreview, verrà visualizzato (1) o meno (0)

  ToolPanelView            Nel caso di RichPreview, indica se nell'area del Treeview vengono visualizzati i parametri (P) interattivi o i raggruppamenti (G) o nulla
  
  LocalizeFromResource     Indica se deve (1) o non deve (0) essere effettuata una localizzazione del report

  ResourcePrefix           Indica un eventuale prefisso da agganciare al nome del report per individuare la propria risorsa di localizzazione.
                           Questo permette di utilizzare lo stesso WRP per versioni differenti del report localizzato

  GetValuesFromHttpPost    Indica se devono (1) o non devono (0) essere letti eventuali parametri passati in post

  OverrideCulture          Indica se devono (1) o non devono (0) essere sovrascritte le impostazioni per la Culture e UICulture nella visualizzazione dei reports
  
  OverridedUICulture       UICulture (formato "<codicelingua2>-<codicepaese2>") per esempio "en-US" per l'inglese (Stati Uniti)

  OverridedCulture         Culture (formato "<codicelingua2>-<codicepaese2>")

  ForceDbPropertiesChange  Serve per applicare una forzatura al cambio delle proprietà della connessione al db. In alcuni casi infatti, 
	                       sembra che la modalità di settaggio standard non funzioni e il report richiede comunque le credenziali per l'accesso ad db una volta lanciato
						   (1) indica se deve essere applicata la forzatura (0) se non deve essere applicata

************************* NOTE PER L'INSTALLAZIONE DI WRP 

*** Pre-requisiti: 

1) .NET Framework 4.0 - Visual Studio 2013

2) Installazione id SAP Crystal Reports (developer version for Microsoft Visual Studio) - runtime 13.x
 
3) Gli eventuali file di risorsa (*.resx) relativi ai reports vanno inseriti nella cartella App_GlobalResources
dell'applicazione seguendo la seguente denominazione: "<NomeReport>.rpt.resx", mentre le traduzioni
dei report relativi devono essere nominati come "<NomeReport>.rpt.<langCode>.resx".

*** Web.config:

1) Il file di configurazione xml (web.config) dell'applicazione è parte integrante della stessa;
per il merge con un pre-esitente file partire sempre, come master, dal nuovo file

2) Attenzione al chiave "TablePrefix", per evitare problemi occorre valorizzare il valore in maiuscolo

*** Soluzioni a problemi noti:

1) Errore lato server: "maximum report processing jobs limit"
Incrementare il "Print Job Limit", il default è 75; nel caso di CR 13 vedere la seguente chiave del registry: HKEY_LOCAL_MACHINE\SOFTWARE\SAP BusinessObjects\Crystal Reports for .NET Framework 4.0\Report Application Server\InprocServer\PrintJobLimit
Fare riferimento al seguente link di SAP:
http://scn.sap.com/community/crystal-reports-for-visual-studio/blog/2014/04/25/what-exactly-is-maximum-report-processing-job-limit-for-crystal-reports

