using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Resources;
using System.IO;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace WRP.ManageResources
{
    public partial class ResourcesGenerator : Form
    {
        public ResourcesGenerator()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // configure the open file dialog
            openFileDialog1.Title = "Add File";
            openFileDialog1.Filter = "Crystal Reports (*.rpt)|*.rpt|" +
                                     "All Files (*.*)|*.*";
            openFileDialog1.FileName = "";
            openFileDialog1.Multiselect = true;

            // return if the user cancels the operation
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            foreach (String filePath in openFileDialog1.FileNames)
            {
                try
                {
                    if (filePath == "")
                        continue;

                    // make sure the file exists before adding
                    // its path to the list of files to be
                    // compressed
                    if (System.IO.File.Exists(filePath) == false)
                        continue;
                    else
                    {
                        if (lstFilePaths.Items.IndexOf(filePath) < 0)
                        {
                            lstFilePaths.Items.Add(filePath);
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }




        }

        private void button2_Click(object sender, EventArgs e)
        {
            // clear the folder path
            txtSaveTo.Text = string.Empty;

            // Show the FolderBrowserDialog.
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtSaveTo.Text = folderBrowserDialog1.SelectedPath;
            }

        }


        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                lstFilePaths.Items.Remove(lstFilePaths.SelectedItem);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            int i = 0;

            // make sure there are files to zip
            if (lstFilePaths.Items.Count < 1)
            {
                MessageBox.Show("There are not files queued for the Resx generation", "Empty File Set");
                return;
            }

            // make sure there is a destination defined
            if (txtSaveTo.Text == string.Empty)
            {
                MessageBox.Show("No destination file has been defined.",
                "Save To Empty");
                return;
            }

            //
            // zip up the files
            try
            {
                string sTargetFolderPath = txtSaveTo.Text;
                string[] filenames = Directory.GetFiles(sTargetFolderPath);

                // ...
                for (i = 0; i < lstFilePaths.Items.Count; i++)
                {
                    string filePath = lstFilePaths.Items[i].ToString();
                    FileInfo fi2 = new FileInfo(filePath);
                    if (fi2.Exists)
                    {
                        // export resx
                        string myFileName = fi2.Name;
                        if (checkBoxAddPrefix.Checked == true)
                        {
                            myFileName = txtPrefix.Text + myFileName;
                        }
                        string newfileResx = sTargetFolderPath + @"\" + myFileName + @".resx";

                        ResXResourceWriter RwX = new ResXResourceWriter(newfileResx);

                        DataTable tmpDt;
                        tmpDt = GetDataTableFromReport(filePath);
                        if (tmpDt != null)
                        {
                            int count = 0;
                            count = tmpDt.Rows.Count;
                            for (int li = 0; li < count; li++)
                            {
                                string key = "";
                                string value = "";
                                key = tmpDt.Rows[li][0].ToString();
                                value = tmpDt.Rows[li][1].ToString();
                                RwX.AddResource(key, value);
                            }
                        }

                        RwX.Generate();
                        RwX.Close();

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Create resources Error");
            }

            MessageBox.Show(string.Format("{0} Resources created.", i));

            Cursor.Current = Cursors.Default;

        }

        private DataTable GetDataTableFromReport(string filePath)
        {

            // DataTable
            DataTable tmpDt = new DataTable("myLoc");
            // Create two column and add to the DataTable.
            tmpDt.Columns.Add("Name", System.Type.GetType("System.String"));
            tmpDt.Columns.Add("Value", System.Type.GetType("System.String"));

            // Work with Report document
            ReportDocument crReportDocument;
            crReportDocument = new ReportDocument();
            //
            try
            {
                // Load
                crReportDocument.Load(filePath, CrystalDecisions.Shared.OpenReportMethod.OpenReportByTempCopy);

                // Loop for text objects
                ReportObjects crReportObjects = crReportDocument.ReportDefinition.ReportObjects;
                foreach (ReportObject crReportObject in crReportObjects)
                {
                    string appoKind = crReportObject.Kind.ToString();
                    if ((crReportObject.Kind == ReportObjectKind.TextObject ||
                        crReportObject.Kind == ReportObjectKind.FieldHeadingObject) &&
                        crReportObject.Name.StartsWith(WRP.ManageResources.Properties.Settings.Default.TranslateTextFieldPrefix))
                    {
                        // Get textobject
                        TextObject myLabel = (TextObject)crReportObject;
                        string reportObjectName = "";
                        string reportObjectValue = "";
                        reportObjectName = myLabel.Name;
                        reportObjectValue = myLabel.Text;
                        // Create DataRow object, and add to DataTable.
                        DataRow locDataRow;
                        locDataRow = tmpDt.NewRow();
                        locDataRow["Name"] = reportObjectName;
                        locDataRow["Value"] = reportObjectValue;
                        tmpDt.Rows.Add(locDataRow);
                    }
                }

                // Loop for formula fields
                FormulaFieldDefinitions crFormulaFieldDefinitions = crReportDocument.DataDefinition.FormulaFields;
                foreach (FormulaFieldDefinition crFormulaFieldDefinition in crFormulaFieldDefinitions)
                {
                    if (crFormulaFieldDefinition.Name.StartsWith(WRP.ManageResources.Properties.Settings.Default.TranslateFormulaFieldPrefix))
                    {
                        // Get textobject
                        FormulaFieldDefinition myLabel = (FormulaFieldDefinition)crFormulaFieldDefinition;
                        string reportObjectName = "";
                        string reportObjectValue = "";
                        reportObjectName = myLabel.Name;
                        reportObjectValue = myLabel.Text;
                        // Create DataRow object, and add to DataTable.
                        DataRow locDataRow;
                        locDataRow = tmpDt.NewRow();
                        locDataRow["Name"] = reportObjectName;
                        locDataRow["Value"] = reportObjectValue;
                        tmpDt.Rows.Add(locDataRow);
                    }
                }

                // Surf the subreports
                Sections crSections;
                ReportDocument crSubreportDocument;
                SubreportObject crSubreportObject;
                string tempSubreportName;
                //
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

                            // open the subreport object
                            tempSubreportName = crSubreportObject.SubreportName;
                            crSubreportDocument = crSubreportObject.OpenSubreport(tempSubreportName);

                            // Loop for text objects
                            ReportObjects crSubReportObjects = crSubreportDocument.ReportDefinition.ReportObjects;
                            foreach (ReportObject crSubReportObject in crSubReportObjects)
                            {
                                if ((crSubReportObject.Kind == ReportObjectKind.TextObject ||
                                    crSubReportObject.Kind == ReportObjectKind.FieldHeadingObject) &&
                                    crSubReportObject.Name.StartsWith(WRP.ManageResources.Properties.Settings.Default.TranslateTextFieldPrefix))
                                {
                                    // Get textobject
                                    TextObject myLabel = (TextObject)crSubReportObject;
                                    string reportObjectName = "";
                                    string reportObjectValue = "";
                                    reportObjectName = myLabel.Name;
                                    reportObjectValue = myLabel.Text;
                                    // Create DataRow object, and add to DataTable.
                                    DataRow locDataRow;
                                    locDataRow = tmpDt.NewRow();
                                    locDataRow["Name"] = reportObjectName;
                                    locDataRow["Value"] = reportObjectValue;
                                    tmpDt.Rows.Add(locDataRow);
                                }
                            }

                            // Loop for formula fields
                            FormulaFieldDefinitions crSubFormulaFieldDefinitions = crSubreportDocument.DataDefinition.FormulaFields;
                            foreach (FormulaFieldDefinition crSubFormulaFieldDefinition in crSubFormulaFieldDefinitions)
                            {
                                if (crSubFormulaFieldDefinition.Name.StartsWith(WRP.ManageResources.Properties.Settings.Default.TranslateFormulaFieldPrefix))
                                {
                                    // Get textobject
                                    FormulaFieldDefinition myLabel = (FormulaFieldDefinition)crSubFormulaFieldDefinition;
                                    string reportObjectName = "";
                                    string reportObjectValue = "";
                                    reportObjectName = myLabel.Name;
                                    reportObjectValue = myLabel.Text;
                                    // Create DataRow object, and add to DataTable.
                                    DataRow locDataRow;
                                    locDataRow = tmpDt.NewRow();
                                    locDataRow["Name"] = reportObjectName;
                                    locDataRow["Value"] = reportObjectValue;
                                    tmpDt.Rows.Add(locDataRow);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw (new Exception("ResourcesGenerator.GetDataTableFromReport: " + ex.Message, ex));
            }

            return tmpDt;
        }

        private void ResourcesGenerator_Load(object sender, EventArgs e)
        {
            // Defaults
            txtSaveTo.Text = WRP.ManageResources.Properties.Settings.Default.SaveToFolder;
            txtPrefix.Text = WRP.ManageResources.Properties.Settings.Default.PrefixNameResx;

        }



    }
}