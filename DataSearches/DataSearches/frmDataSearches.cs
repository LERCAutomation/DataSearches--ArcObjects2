using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GeoDatabaseUI;

using HLSearchesToolConfig;
using HLFileFunctions;
using HLStringFunctions; 
using HLArcMapModule;

using System.Data.OleDb;

namespace DataSearches
{
    public partial class frmDataSearches : Form
    {
        SearchesToolConfig myConfig;
        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        ArcMapFunctions myArcMapFuncs;
        string strUserID;
        string strTempFile;

        bool blOpenForm; // this tracks all the way through initialisation whether the form should open.
        bool blFormHasOpened; // This informs all controls whether the form has opened.

        public frmDataSearches()
        {
            blOpenForm = true;
            blFormHasOpened = false;
            InitializeComponent();
            myConfig = new SearchesToolConfig();
            myFileFuncs = new FileFunctions();
            IApplication pApp = ArcMap.Application;
            myArcMapFuncs = new ArcMapFunctions(pApp);
            myStringFuncs = new StringFunctions();
            

            // Get the relevant from the Config file.
            if (myConfig.GetFoundXML() == false)
            {
                MessageBox.Show("XML file not found; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }
            else if (myConfig.GetLoadedXML() == false)
            {
                MessageBox.Show("Error loading XML File; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            if (!blOpenForm)
            {
                    Load += (s, e) => Close();
                    return;
            }

            // Delete any temporary shapefile
            strUserID = Environment.UserName;
            string strSaveRootDir = myConfig.GetSaveRootDir();
            string strTempFolder = strSaveRootDir + @"\Temp";
            if (!Directory.Exists(strTempFolder)) // Create new, hidden directory.
            {
                DirectoryInfo di = Directory.CreateDirectory(strTempFolder);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
            strTempFile = strTempFolder + @"\TempShapes_" + strUserID + ".shp";
            if (myFileFuncs.FileExists(strTempFile))
            {
                myArcMapFuncs.DeleteFeatureclass(strTempFile); // This may not work but exception is handled.
            }


            if (blOpenForm)
            {
                // We've found the XML and loaded it successfully. Fill in the form.
                List<string> AllLayers = myConfig.GetMapLayers(); // All possible layers by name
                List<string> AllDisplayLayers = myConfig.GetMapNames(); // All possible layers by display name
                List<bool> blLoadWarnings = myConfig.GetMapLoadWarnings(); // A list telling us whether to warn users if layer not present
                List<bool> blPreselectLayers = myConfig.GetMapPreselectLayers(); // A list telling us which layers to preselect in the form
                List<string> OpenLayers = new List<string>(); // The open layers by name
                List<bool> PreselectLayers = new List<bool>(); // The preselect options of the open layers
                List<string> ClosedLayers = new List<string>(); // The closed layers by name

                int i = 0;
                foreach (string aLayer in AllDisplayLayers)
                {
                    if (myArcMapFuncs.LayerExists(aLayer))
                    {
                        OpenLayers.Add(AllLayers[i]);
                        PreselectLayers.Add(blPreselectLayers[i]);
                    }
                    else
                    {
                        if (blLoadWarnings[i] == true) // Only add if the user wants to be warned of this one.
                            ClosedLayers.Add(aLayer);
                    }
                    i++;
                }

                // Add the available layers to the form.
                lstLayers.Items.AddRange(OpenLayers.ToArray());

                // Highlight the preselected ones
                i = 0;
                foreach (string aLayer in OpenLayers)
                {
                    lstLayers.SetSelected(i, PreselectLayers[i]);
                    i++;
                }


                // Fill in the rest of the form.
                txtBufferSize.Text = myConfig.GetDefaultBufferSize().ToString();
                cmbUnits.Items.AddRange(myConfig.GetBufferUnitOptionsDisplay().ToArray());
                cmbUnits.SelectedIndex = myConfig.GetDefaultBufferUnit() - 1;
                cmbAddLayers.Items.AddRange(myConfig.GetAddSelectedLayerOptions().ToArray());
                if (myConfig.GetDefaultAddSelectedLayerOption() != -1)
                    cmbAddLayers.SelectedIndex = myConfig.GetDefaultAddSelectedLayerOption() - 1;
                cmbLabels.Items.AddRange(myConfig.GetOverwriteLabelOptions().ToArray());
                if (myConfig.GetDefaultOverwriteLabelsOption() != -1)
                    cmbLabels.SelectedIndex = myConfig.GetDefaultOverwriteLabelsOption() - 1;

                cmbCombinedSites.Items.AddRange(myConfig.GetCombinedSitesTableOptions().ToArray());
                if (myConfig.GetDefaultCombinedSitesTable() != -1)
                    cmbCombinedSites.SelectedIndex = myConfig.GetDefaultCombinedSitesTable() - 1;

                chkClearLog.Checked = myConfig.GetDefaultClearLogFile();

                // Hide controls that were not requested
                if (myConfig.GetDefaultAddSelectedLayerOption() == -1)
                {
                    cmbAddLayers.Hide();
                    label5.Hide();
                }
                else
                {
                    cmbAddLayers.Show();
                    label5.Show();
                }
                if (myConfig.GetDefaultOverwriteLabelsOption() == -1 || myConfig.GetDefaultAddSelectedLayerOption() == -1)
                {
                    cmbLabels.Hide();
                    label7.Hide();
                }
                else 
                {
                    cmbLabels.Show();
                    label7.Show();
                }
                if (myConfig.GetDefaultCombinedSitesTable() == -1)
                {
                    cmbCombinedSites.Hide();
                    label8.Hide();
                }
                else
                {
                    cmbCombinedSites.Show();
                    label8.Show();
                }

                // Warn the user of closed layers.
                if (ClosedLayers.Count > 0)
                {
                    string strWarning = "";
                    if (ClosedLayers.Count == 1)
                        strWarning = "Warning. The layer '" + ClosedLayers[0] + "' is not loaded.";
                    else
                    {
                        strWarning = "Warning: " + ClosedLayers.Count.ToString() + " layers are not loaded, including ";
                        foreach (string strLayer in ClosedLayers)
                            strWarning = strWarning + strLayer + "; ";
                    }
                    MessageBox.Show(strWarning);
                }
                blFormHasOpened = true;
            }
            else // Something has gone wrong during initialisation; don't load form.
            {
                Load += (s, e) => Close();
                return;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            // Obtain all the relevant information.
            string strReference = txtSearch.Text; // Search reference
            string strSiteName = txtLocation.Text; // Associated location

            // The selected layers
            List<string> SelectedLayers = new List<string>();
            foreach (string strSelectedItem in lstLayers.SelectedItems)
            {
                SelectedLayers.Add(strSelectedItem);
            }


            bool blClearLogFile = chkClearLog.Checked; // Clear log file
            string strBufferSize = txtBufferSize.Text; // Buffer size 
            int intBufferUnitIndex = cmbUnits.SelectedIndex; // Index of the selected item in Buffer Units combobox
            // What does the unit index translate to?
            string strBufferUnitProcess = myConfig.GetBufferUnitOptionsProcess()[intBufferUnitIndex]; // Unit to be used in process (because of American spelling)
            string strBufferUnitText = cmbUnits.Text; // Unit to be used in reporting.
            string strBufferUnitShort = myConfig.GetBufferUnitOptionsShort()[intBufferUnitIndex]; // Unit to be used in file naming (abbreviation)

            string strAddSelected;
            if (cmbAddLayers.CanSelect)
                strAddSelected = cmbAddLayers.Text; // Add selected layers
            else
                strAddSelected = "No";

            string strOverwriteLabels;
            if (cmbLabels.CanSelect)
                strOverwriteLabels = cmbLabels.Text; // Overwrite labels
            else
                strOverwriteLabels = "No";

            bool blResetGroups = false;
            if (cmbLabels.Text.ToLower().Contains("group"))
                blResetGroups = true;

            bool blCombinedTable;
            bool blCombinedTableOverwrite;
            if (cmbCombinedSites.CanSelect)
                if (cmbCombinedSites.Text.ToLower().Contains("none"))
                {
                    blCombinedTable = false;
                    blCombinedTableOverwrite = false;
                }
                else if (cmbCombinedSites.Text.ToLower().Contains("append"))
                {
                    blCombinedTable = true;
                    blCombinedTableOverwrite = false;
                }
                else
                {
                    blCombinedTable = true;
                    blCombinedTableOverwrite = true;
                }
            else
            {
                blCombinedTable = false;
                blCombinedTableOverwrite = false;
            }

            // Check that the user has entered all information correctly.
            if (strReference == "")
            {
                MessageBox.Show("Please enter a search reference");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            if (strSiteName == "")
            {
                MessageBox.Show("Please enter a location");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            if (SelectedLayers.Count == 0)
            {
                MessageBox.Show("Please select at least one layer to search");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            if (strBufferSize == "")
            {
                MessageBox.Show("Please enter a buffer size");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            double i;
            bool blTest = Double.TryParse(strBufferSize, out i);
            if (!blTest || i < 0) // User either entered text or a negative number
            {
                MessageBox.Show("Please enter a positive number for the buffer size");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // All entered information is correct. Now set up the process.
            string strReplaceChar = myConfig.GetReplaceChar();

            // Is the replacement character valid?
            bool blValid = myStringFuncs.isValid(strReplaceChar);
            if (!blValid)
            {
                MessageBox.Show("The replacement character '" + strReplaceChar + "' is not valid. Process aborted");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // fix any illegal characters in the site name string
            strSiteName = myStringFuncs.StripIllegals(strSiteName,strReplaceChar);

            // Create the shortref from the search reference.
            string strShortRef = strReference.Replace("/", strReplaceChar);
            // Get rid of any characters in the shortref.
            strShortRef = myStringFuncs.KeepNumbersAndSpaces(strShortRef, strReplaceChar);
            // Find the subref part of this reference.
            string strSubref = myStringFuncs.GetSubref(strShortRef, strReplaceChar);         

            // Now do any replacements that are necessary in the file names.
            string strSaveRootDir = myConfig.GetSaveRootDir();
            string strSaveFolder = myConfig.GetSaveFolder();
            string strGISFolder = myConfig.GetGISFolder();
            string strLogFileName = myConfig.GetLogFileName();

            string strUserID = Environment.UserName;
            string strTempFolder = strSaveRootDir + @"\Temp";
            string strTempFile = strTempFolder + @"\TempShapes_" + strUserID + ".shp";

            strSaveFolder = strSaveFolder.Replace("%ref%", strReference);
            strSaveFolder = strSaveFolder.Replace("%shortref%", strShortRef);
            strSaveFolder = strSaveFolder.Replace("%subref%", strSubref);
            strSaveFolder = strSaveFolder.Replace("%sitename%", strSiteName);

            strGISFolder = strGISFolder.Replace("%ref%", strReference);
            strGISFolder = strGISFolder.Replace("%shortref%", strShortRef);
            strGISFolder = strGISFolder.Replace("%subref%", strSubref);
            strGISFolder = strGISFolder.Replace("%sitename%", strSiteName);

            strLogFileName = strLogFileName.Replace("%ref%", strReference);
            strLogFileName = strLogFileName.Replace("%shortref%", strShortRef);
            strLogFileName = strLogFileName.Replace("%subref%", strSubref);
            strLogFileName = strLogFileName.Replace("%sitename%", strSiteName);

            // Remove any illegal characters from the names.
            strSaveFolder = myStringFuncs.StripIllegals(strSaveFolder, strReplaceChar);
            strGISFolder = myStringFuncs.StripIllegals(strGISFolder, strReplaceChar);
            strLogFileName = myStringFuncs.StripIllegals(strLogFileName, strReplaceChar, true);

            // Trim any trailing spaces since directory functions don't deal with them and it causes a crash.
            strSaveFolder = strSaveFolder.Trim();
            

            // Create directories as required.
            if (!myFileFuncs.DirExists(strSaveRootDir))
            {
                try
                {
                    Directory.CreateDirectory(strSaveRootDir);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + strSaveRootDir + ". System error: " + ex.Message);
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
            }
            if (strSaveFolder != "")
                strSaveFolder = strSaveRootDir + @"\" + strSaveFolder;
            else
                strSaveFolder = strSaveRootDir;

            if (strGISFolder != "")
                strGISFolder = strSaveFolder + @"\" + strGISFolder;
            else
                strGISFolder = strSaveFolder;

            if (!myFileFuncs.DirExists(strSaveFolder))
            {
                try
                {
                    Directory.CreateDirectory(strSaveFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + strSaveFolder + ". System error: " + ex.Message);
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
            }
            if (!myFileFuncs.DirExists(strGISFolder))
            {
                try
                {
                    Directory.CreateDirectory(strGISFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory " + strGISFolder + ". System error: " + ex.Message);
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
            }

            // All output directories are present.

            // Create logfile if necessary
            string strLogFile = strGISFolder + @"\" + strLogFileName;
            if (myFileFuncs.FileExists(strLogFile) && chkClearLog.Checked == true)
            {
                try
                {
                    File.Delete(strLogFile);
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Cannot clear log file " + strLogFile + ". Please make sure this file is not open in another window. " +
                        "System error: " + ex.Message);
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
                
            }

            // Write the first line to the log file.
            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");
            myFileFuncs.WriteLine(strLogFile, "Processing search " + strReference);
            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");

            myFileFuncs.WriteLine(strLogFile, "Parameters are as follows:");
            myFileFuncs.WriteLine(strLogFile, "Buffer distance: " + strBufferSize + " " + strBufferUnitText);
            myFileFuncs.WriteLine(strLogFile, "Output location: " + strSaveFolder);
            string strLayerString = "";
            foreach (string aLayer in SelectedLayers)
                strLayerString = strLayerString + aLayer + ", ";
            myFileFuncs.WriteLine(strLogFile, "Layers included: " + strLayerString.Substring(0, strLayerString.Length - 2));

            string strSearchLayerBase = myConfig.GetSearchLayer();
            List<string> strSearchLayerExtensions = myConfig.GetSearchLayerExtensions();
            string strSearchColumn = myConfig.GetSearchColumn();

            // Get the search reference object from the SearchLayer.
            string strQuery = strSearchColumn + " = '" + strReference + "'"; // Create the query. Example: "TERM = 'City'"

            // The output file for the buffer is a shapefile in the root save directory.
            string strSaveBuffer = strBufferSize;
            if (strSaveBuffer.Contains('.'))
                strSaveBuffer = strSaveBuffer.Replace('.', '_');
            string strLayerName = "Buffer_" + strSubref + "_" + strSaveBuffer + strBufferUnitShort;
            string strOutputFile = strGISFolder + "\\" + strLayerName + ".shp";
            string strSubGroupLayerName = strSubref + "_" + strBufferSize + strBufferUnitShort;
            string strBufferFields = myConfig.GetAggregateColumns();

            // Find the search feature.
            int aCount = 0;
            string strTargetLayer = null;

            foreach (string strExtension in strSearchLayerExtensions)
            {
                string strSearchLayer = strSearchLayerBase + strExtension;

                // Get the feature if it exists. Only search existing layers.
                if (myArcMapFuncs.LayerExists(strSearchLayer))
                {

                    //SearchLayerList.Add(strSearchLayer);
                    myFileFuncs.WriteLine(strLogFile, "Searching layer " + strSearchLayer + " for reference " + strReference);
                    IFeature pTestFeature = myArcMapFuncs.GetFeatureFromLayer(strSearchLayer, strQuery);
                    if (pTestFeature != null)
                    {
                        if (aCount == 0)
                        {
                            strTargetLayer = strSearchLayer;
                            aCount = aCount + 1;
                            myFileFuncs.WriteLine(strLogFile, "Found feature with reference " + strReference + " in layer " + strSearchLayer);
                        }
                        else
                            aCount = aCount + 1;
                    }
                }
            }

            if (aCount == 0)
            {
                MessageBox.Show("The feature with reference " + strReference + " was not found in any of the search layers");
                myFileFuncs.WriteLine(strLogFile, "Feature with reference " + strReference + " was not found in any of the search layers");
                myFileFuncs.WriteLine(strLogFile, "Process aborted");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }
            if (aCount > 1)
            {
                MessageBox.Show("There were " + aCount.ToString() + " features with the same reference " + strReference + " in the search layers; Process aborted");
                myFileFuncs.WriteLine(strLogFile, "There were " + aCount.ToString() + " features with the same reference " + strReference + " in the search layers");
                myFileFuncs.WriteLine(strLogFile, "Process aborted");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }
            
            // We have found the feature and are ready to go. Pause drawing.
            myArcMapFuncs.ToggleDrawing(false);
            myArcMapFuncs.ToggleTOC();

            // Select the feature in the map.
            myArcMapFuncs.SelectLayerByAttributes(strTargetLayer, strQuery);

            // Get the required information for the aggregate columns.

            // Create a buffer around the feature and save into a new file.
            // convert to the correct unit.
           
            string strBufferString = strBufferSize + " " + strBufferUnitProcess;
            if (strBufferSize == "0") strBufferString = "0.01 Meters";  // Safeguard for zero buffer size; Select a tiny buffer to allow
                                                                       // correct legending (expects a polygon).


            myFileFuncs.WriteLine(strLogFile, "Buffering feature from " + strTargetLayer + " with a distance of " + strBufferSize + " " + strBufferUnitText);
            bool blResult = myArcMapFuncs.BufferFeature(strTargetLayer, strOutputFile, strBufferString, strBufferFields, true, true);

            if (blResult == true)
            {
                myFileFuncs.WriteLine(strLogFile, "Buffering complete");
            }
            else
            {
                MessageBox.Show("Error during feature buffering. Process aborted");
                myFileFuncs.WriteLine(strLogFile, "Error during feature buffering. Process aborted");
                this.BringToFront();
                this.Cursor = Cursors.Default;
                myArcMapFuncs.ToggleDrawing(true);
                myArcMapFuncs.ToggleTOC();
                return;
            }

            // Add aggregate columns to the buffer.

            // Clear the selected features.
            myArcMapFuncs.ClearSelectedMapFeatures(strTargetLayer);

            // Is the buffer in the map? If not, add it.
            if (!myArcMapFuncs.LayerExists(strLayerName))
            {
                myArcMapFuncs.AddFeatureLayerFromString(strOutputFile);
            }

            // Add the buffer to the group layer. Zoom to the buffer.
            myArcMapFuncs.MoveToSubGroupLayer(strShortRef, strSubGroupLayerName, myArcMapFuncs.GetLayer(strLayerName));
            
            // Change the symbology of the buffer.
            string strLayerDir = myConfig.GetLayerDir();
            string strLayerFile = strLayerDir + @"\" + myConfig.GetBufferLayer();
            myArcMapFuncs.ChangeLegend(strLayerName, strLayerFile);
            myFileFuncs.WriteLine(strLogFile, "Buffer added to display");

            // go through each of the requested layers and carry out the relevant analysis. 
            List<string> strLayerNames = myConfig.GetMapLayers();
            List<string> strDisplayNames = myConfig.GetMapNames();
            List<string> strPrefixes = myConfig.GetMapPrefixes();
            List<string> strSuffixes = myConfig.GetMapSuffixes();
            List<string> strColumnList = myConfig.GetMapColumns();
            List<string> strGroupColumnList = myConfig.GetMapGroupByColumns();
            List<string> strStatsColumnList = myConfig.GetMapStatisticsColumns();
            List<string> strOrderColumnList = myConfig.GetMapOrderByColumns();
            List<string> strCriteriaList = myConfig.GetMapCriteria();
            List<bool> blIncludeDistances = myConfig.getMapIncludeDistances();
            List<bool> blIncludeRadii = myConfig.getMapIncludeRadii();
            List<string> strKeyColumns = myConfig.GetMapKeyColumns();
            List<string> strFormats = myConfig.GetMapFormats();
            List<bool> blKeepLayers = myConfig.GetMapKeepLayers();
            List<string> strDisplayLayerFiles = myConfig.GetMapLayerFiles();
            List<bool> blOverwriteLabelDefaults = myConfig.GetMapOverwriteLabels();
            List<string> strLabelColumns = myConfig.GetMapLabelColumns();
            List<string> strLabelClauses = myConfig.GetMapLabelClauses();
            List<string> strCombinedSitesColumnList = myConfig.GetMapCombinedSitesColumns();
            List<string> strCombinedSitesGroupColumnList = myConfig.GetMapCombinedSitesGroupByColumns();
            List<string> strCombinedSitesStatsColumnList = myConfig.GetMapCombinedSitesStatsColumns();
            List<string> strCombinedSitesOrderColumnList = myConfig.GetMapCombinedSitesOrderByColumns();
            //List<string> strCombinedSitesCriteriaList = myConfig.GetMapCombinedSitesCriteria();

            // Start the combined sites table before we do any analysis.
            //string strCombinedTable = myConfig.GetCombinedSitesTableName();
            string strCombinedFormat = myConfig.GetCombinedSitesTableFormat();
            string strCombinedSuffix = myConfig.GetCombinedSitesTableSuffix();
            string strCombinedTable = strGISFolder + @"\" + strSubref + strCombinedSuffix + "." + strCombinedFormat;
            if (blCombinedTable)
            {
                string strCombinedSitesHeader = myConfig.GetCombinedSitesTableColumns();
                // Start the table if overwrite has been selected, or if the table doesn't exist and Append has been selected.
                if (blCombinedTableOverwrite || (!myFileFuncs.FileExists(strCombinedTable) && !blCombinedTableOverwrite))
                {
                    blResult = myArcMapFuncs.WriteEmptyCSV(strCombinedTable, strCombinedSitesHeader);
                    if (!blResult)
                    {
                        MessageBox.Show("Error writing to combined sites table. Process aborted");
                        myFileFuncs.WriteLine(strLogFile, "Error writing to combined sites table. Process aborted");
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        myArcMapFuncs.ToggleDrawing(true);
                        myArcMapFuncs.ToggleTOC();
                        return;
                    }
                    myFileFuncs.WriteLine(strLogFile, "Combined sites table started");
                }
            }

            // Now go through the layers.

            // Get any groups and initialise required layers.
            List<string> liGroupNames = new List<string>();
            List<int> liGroupLabels = new List<int>();
            if (blResetGroups)
            {
                liGroupNames = myStringFuncs.ExtractGroups(SelectedLayers);
                foreach (string strGroupName in liGroupNames)
                {
                    liGroupLabels.Add(1); // each group has its own label counter.
                }
            }

            int intStartLabel = 1; // Keep track of the label numbers if there are no groups.
            int intMaxLabel = 1;
            foreach (string aLayer in SelectedLayers)
            {
                myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");
                myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");
                // Get all the settings relevant to this layer.
                int intIndex = strLayerNames.IndexOf(aLayer); // Finds the first occurrence. 
                string strDisplayName = strDisplayNames[intIndex];
                string strPrefix = strPrefixes[intIndex];
                string strSuffix = strSuffixes[intIndex];
                string strColumns = strColumnList[intIndex]; // Note there could be multiple columns.
                string strGroupColumns = strGroupColumnList[intIndex];
                string strStatsColumns = strStatsColumnList[intIndex];
                string strOrderColumns = strOrderColumnList[intIndex];
                string strCriteria = strCriteriaList[intIndex];
                bool blIncludeDistance = blIncludeDistances[intIndex];
                bool blIncludeRadius = blIncludeRadii[intIndex];
                
                string strKeyColumn = strKeyColumns[intIndex];
                string strFormat = strFormats[intIndex];
                bool blKeepLayer = blKeepLayers[intIndex];
                string strDisplayLayer = strDisplayLayerFiles[intIndex];
                bool blOverwriteLabelDefault = blOverwriteLabelDefaults[intIndex];
                string strLabelColumn = strLabelColumns[intIndex];
                string strLabelClause = strLabelClauses[intIndex];
                string strCombinedSitesColumns = strCombinedSitesColumnList[intIndex];
                string strCombinedSitesGroupColumns = strCombinedSitesGroupColumnList[intIndex];
                string strCombinedSitesStatsColumns = strCombinedSitesStatsColumnList[intIndex];
                string strCombinedSitesOrderColumns = strCombinedSitesOrderColumnList[intIndex];
                //string strCombinedSitesCriteria = strCombinedSitesCriteriaList[intIndex];


                strStatsColumns = myStringFuncs.AlignStatsColumns(strColumns, strStatsColumns, strGroupColumns);
                //if (blIncludeDistance && !strColumns.Contains("Distance") && !strGroupColumns.Contains("Distance"))
                //    strColumns = strColumns + ",Distance"; // Distance comes after grouping and hence should not be included in the stats columns.
                //if (blIncludeRadius && !strColumns.Contains("Radius") && !strGroupColumns.Contains("Radius"))
                //    strColumns = strColumns + ",Radius"; // as for Distance column, it comes after the grouping.
                strCombinedSitesStatsColumns = myStringFuncs.AlignStatsColumns(strCombinedSitesColumns, strCombinedSitesStatsColumns, strCombinedSitesGroupColumns);

                // Create relevant output name. Note this is done whether or not the layer is eventually kept.
                string strShapeLayerName = strPrefix + "_" + strSubref; // Use the prefix as the layer name.
                string strShapeOutputName = strGISFolder + @"\" + strShapeLayerName + ".shp"; // output shapefile / feature class name
                string strTableOutputName = strGISFolder + @"\" + strSubref + strSuffix + "." + strFormat; // output table name

                myFileFuncs.WriteLine(strLogFile, "Starting analysis for " + aLayer);

                // Do the selection. 
                myFileFuncs.WriteLine(strLogFile, "Selecting features on layer " + strDisplayName + " using selected feature(s) in layer " + strLayerName);
                myArcMapFuncs.SelectLayerByLocation(strDisplayName, strLayerName);

                // Refine the selection if required.
                if (myArcMapFuncs.CountSelectedLayerFeatures(strDisplayName) > 0 && strCriteria != "")
                {
                    myFileFuncs.WriteLine(strLogFile, "Refining selection on " + strDisplayName + " with criteria " + strCriteria);
                    blResult = myArcMapFuncs.SelectLayerByAttributes(strDisplayName, strCriteria, "SUBSET_SELECTION");
                    if (!blResult)
                    {
                        MessageBox.Show("Error selecting layer " + strDisplayName + " with criteria " + strCriteria + ". Please check syntax and column names (case sensitive)");
                        myFileFuncs.WriteLine(strLogFile, "Error refining selection on layer " + strDisplayName + " with criteria " + strCriteria + ". Please check syntax and column names (case sensitive)");
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        myArcMapFuncs.ToggleDrawing(true);
                        myArcMapFuncs.ToggleTOC();
                        return;
                    }
                    myFileFuncs.WriteLine(strLogFile, "Selection on " + strDisplayName + " refined");
                }

                // Write out the results - to shapefile first. Include distance if required.
                // Function takes account of output, group by and statistics fields.

                if (myArcMapFuncs.CountSelectedLayerFeatures(strDisplayName) > 0)
                {
                    myFileFuncs.WriteLine(strLogFile, myArcMapFuncs.CountSelectedLayerFeatures(strDisplayName).ToString() + " feature(s) found for " + strDisplayName + ". Processing");
                    // Firstly take a copy of the full selection in a temporary file; This will be used to do the summaries on.
                    string strTempMaster = "TempMaster" + strUserID;
                    string strTempMasterOutput = strTempFolder + @"\" + strTempMaster + ".shp";
                    blResult = myArcMapFuncs.CopyFeatures(strDisplayName, strTempMasterOutput);
                    if (!blResult)
                    {
                        MessageBox.Show("Cannot copy selection from " + strDisplayName + " to " + strTempMasterOutput + ". Please ensure this file is not open elsewhere");
                        myFileFuncs.WriteLine(strLogFile, "Cannot copy selection from " + strDisplayName + " to " + strTempMasterOutput + ". Please ensure this file is not open elsewhere");
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        myArcMapFuncs.ToggleDrawing(true);
                        myArcMapFuncs.ToggleTOC();
                        return;
                    }
                    

                    // Check if the map label field exists. Create if necessary. 
                    string strGroupName = myStringFuncs.GetGroupName(aLayer);
                    bool blNewLabelField = false;
                    if (!myArcMapFuncs.FieldExists(strTempMasterOutput, strLabelColumn) && strAddSelected.ToLower().Contains("with") && strLabelColumn != "")
                    {
                        // If not, create it and label.
                        myArcMapFuncs.AddField(strTempMasterOutput, strLabelColumn, esriFieldType.esriFieldTypeInteger, 10);
                        blNewLabelField = true;
                    }
                    
                    // Add labels as required
                    if (blNewLabelField || (strOverwriteLabels.ToLower() != "no" && blOverwriteLabelDefault && strAddSelected.ToLower().Contains("with") && strLabelColumn != ""))
                    // Either we  have a new label field, or we want to overwrite the labels and are allowed to.
                    {
                        // Add relevant labels. 
                        if (strOverwriteLabels.ToLower().Contains("layer")) // Reset each layer to 1.
                        {
                            myFileFuncs.WriteLine(strLogFile, "Resetting label counter");
                            intStartLabel = 1;
                            myArcMapFuncs.AddIncrementalNumbers(strTempMasterOutput, strLabelColumn, strKeyColumn, intStartLabel);
                            myFileFuncs.WriteLine(strLogFile, "Map labels added");
                        }

                        else if (blResetGroups && strGroupName != "")
                        {
                            // Increment within but reset between groups. Note all group labels are already initialised as 1.
                            // Only triggered if a group name has been found.

                            int intGroupIndex = liGroupNames.IndexOf(strGroupName);
                            int intGroupLabel = liGroupLabels[intGroupIndex];
                            intGroupLabel = myArcMapFuncs.AddIncrementalNumbers(strTempMasterOutput, strLabelColumn, strKeyColumn, intGroupLabel);
                            intGroupLabel++;
                            liGroupLabels[intGroupIndex] = intGroupLabel; // Store the new max label.
                            myFileFuncs.WriteLine(strLogFile, "Map labels added");
                        }
                        else
                        {
                            // There is no group or groups are ignored, or we are not resetting. Use the existing max label number.
                            intStartLabel = intMaxLabel;

                            intMaxLabel = myArcMapFuncs.AddIncrementalNumbers(strTempMasterOutput, strLabelColumn, strKeyColumn, intStartLabel);
                            intMaxLabel++; // the new start label for incremental labeling
                            myFileFuncs.WriteLine(strLogFile, "Map labels added");
                        }
                    }

                    string strTempOutput = "TempOutput" + strUserID;
                    string strTempShapeOutput = strTempFolder + @"\" + strTempOutput + ".shp";
                    myFileFuncs.WriteLine(strLogFile, "Extracting summary information from " + strDisplayName);
                    string strRadius = "none";
                    if (blIncludeRadius) strRadius = strBufferSize + " " + strBufferUnitText; // Only include radius if requested.
                    myArcMapFuncs.ExportSelectionToShapefile(strTempMaster, strTempShapeOutput, strColumns, strTempFile, strGroupColumns,
                        strStatsColumns, blIncludeDistance, strRadius, strTargetLayer);

                    // Write out the results to table as appropriate.
                    bool blIncHeaders = false;
                    if (strFormat.ToLower() == "csv") blIncHeaders = true;

                    myFileFuncs.WriteLine(strLogFile, "Writing summary information for " + strDisplayName + " to " + strTableOutputName);
                    // 29/11/2016 note no longer includes strCriteria in the export as taken care of above. 
                    int intLineCount = myArcMapFuncs.CopyToCSV(strTempShapeOutput, strTableOutputName, strColumns, strOrderColumns, true, false, !blIncHeaders);
                    myFileFuncs.WriteLine(strLogFile, intLineCount.ToString() + " line(s) written for " + strDisplayName);

                    // Copy to permanent layer as appropriate
                    if (blKeepLayer && strAddSelected.ToLower() != "no")
                    {
                        myFileFuncs.WriteLine(strLogFile, "Copying selected GIS features from " + strDisplayName + " to " + strShapeOutputName);
                        myArcMapFuncs.CopyFeatures(strTempMaster, strShapeOutputName);
                        myArcMapFuncs.MoveToSubGroupLayer(strShortRef, strSubGroupLayerName, myArcMapFuncs.GetLayer(strShapeLayerName));
                        if (strDisplayLayer != "")
                        {
                            string strDisplayLayerFile = strLayerDir + @"\" + strDisplayLayer;
                            myArcMapFuncs.ChangeLegend(strShapeLayerName, strDisplayLayerFile);
                        }
                        myFileFuncs.WriteLine(strLogFile, "Output " + strShapeLayerName + " added to display");
                        if (strAddSelected.ToLower().Contains("with "))
                        {
                            // Translate the label string.
                            if (strLabelClause != "")
                            {
                                List<string> strLabelOptions = strLabelClause.Split('$').ToList();
                                string strFont = strLabelOptions[0].Split(':')[1];
                                double dblSize = double.Parse(strLabelOptions[1].Split(':')[1]); // Needs error trapping
                                int intRed = int.Parse(strLabelOptions[2].Split(':')[1]); // Needs error trapping
                                int intGreen = int.Parse(strLabelOptions[3].Split(':')[1]);
                                int intBlue = int.Parse(strLabelOptions[4].Split(':')[1]);
                                string strOverlap = strLabelOptions[5].Split(':')[1];
                                myArcMapFuncs.AnnotateLayer(strShapeLayerName, "[" + strLabelColumn + "]", strFont, dblSize,
                                    intRed, intGreen, intBlue, strOverlap);
                                myFileFuncs.WriteLine(strLogFile, "Labels added to output " + strShapeLayerName);
                            }
                            else if (strLabelColumn != "")
                            {
                                myArcMapFuncs.AnnotateLayer(strShapeLayerName, "[" + strLabelColumn + "]");
                                myFileFuncs.WriteLine(strLogFile, "Labels added to output " + strShapeLayerName);
                            }
                        }
                    }
                    myArcMapFuncs.RemoveLayer(strTempOutput);
                    myArcMapFuncs.DeleteFeatureclass(strTempShapeOutput);
                    

                    // Add to combined sites table as appropriate
                    // Function to take account of group by, order by and statistics fields.
                    if (strCombinedSitesColumns != "" && blCombinedTable)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Extracting summary output for combined sites table");
                        myArcMapFuncs.ExportSelectionToShapefile(strTempMaster, strTempShapeOutput, strCombinedSitesColumns, strTempFile,
                            strCombinedSitesGroupColumns, strCombinedSitesStatsColumns, blIncludeDistance, strRadius, strTargetLayer );

                        // This needs changing to take account of the 'tag' field. Will need a new function that takes account
                        // of the combined sites header columns.
                        myFileFuncs.WriteLine(strLogFile, "Writing summary output to combined sites table " + strCombinedTable);
                        intLineCount = myArcMapFuncs.CopyToCSV(strTempShapeOutput, strCombinedTable, strCombinedSitesColumns, strCombinedSitesOrderColumns, true, true);
                        myFileFuncs.WriteLine(strLogFile, "Summary output written. " + intLineCount.ToString() + " row(s) added for " + strDisplayName);
                        myArcMapFuncs.RemoveLayer(strTempOutput);
                        myArcMapFuncs.DeleteFeatureclass(strTempShapeOutput);
              

                    }
                    myArcMapFuncs.RemoveLayer(strTempMaster);
                    myArcMapFuncs.DeleteFeatureclass(strTempMasterOutput);

                    // Clear the selection.
                    myArcMapFuncs.ClearSelectedMapFeatures(strDisplayName);
                    myFileFuncs.WriteLine(strLogFile, "Analysis complete for " + aLayer);
                }
                else
                {
                    myFileFuncs.WriteLine(strLogFile, "No features selected for " + aLayer);
                }


            }

            // All done, bring to front etc. 
            myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");
            myFileFuncs.WriteLine(strLogFile, "Process complete");
            myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");

            myArcMapFuncs.ToggleTOC();
            myArcMapFuncs.ToggleDrawing(true);
            if (strBufferSize != "0")
                myArcMapFuncs.ZoomToLayer(strLayerName);


            this.Cursor = Cursors.Default;
            DialogResult dlResult = MessageBox.Show("Process complete. Do you wish to close the form?", "Data Searches", MessageBoxButtons.YesNo);
            if (dlResult == System.Windows.Forms.DialogResult.Yes)
                this.Close();
            else this.BringToFront();

            Process.Start("notepad.exe", strLogFile);

        }
              

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            
            if (txtSearch.Text != "" && txtSearch.Text.Length > 11) // Only fire it when it looks like we have a complete reference.
            {
                
                string strAccessConn = "Provider='Microsoft.Jet.OLEDB.4.0';data source='" + myConfig.GetDatabase() + "'";
                string strQuery = "SELECT " + myConfig.GetSiteColumn() + " from Enquiries WHERE " + myConfig.GetRefColumn() + " = " + '"' + txtSearch.Text + '"';
                string strLocation = "";
                OleDbConnection myAccessConn = null;
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to create a database connection. System error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DataSet myDataSet = new DataSet();
                try
                {

                    OleDbCommand myAccessCommand = new OleDbCommand(strQuery, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    myAccessConn.Open();
                    myDataAdapter.Fill(myDataSet, "Enquiries");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the database. System error:" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    myAccessConn.Close();
                }

                DataRowCollection myRS = myDataSet.Tables["Enquiries"].Rows;
                foreach (DataRow aRow in myRS) // Really there should only be one. We can check for this.
                {
                    // Get the location name
                    strLocation = aRow[0].ToString();
                }

                if (strLocation != "")
                {
                    // The location is known. Fill it in and do not allow editing.
                    txtLocation.Text = strLocation;
                    txtLocation.Enabled = false;
                }
                else
                {
                    // The location is not known. Allow user to enter it.
                    txtLocation.Text = "";
                    txtLocation.Enabled = true;
                    // Should we allow an update??
                }

            }
            else
            {
                // There is no search reference. Disable the Location text box.
                txtLocation.Text = "";
                txtLocation.Enabled = false;
            }
        }

        private void cmbAddLayers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbAddLayers.Text == "Yes - With labels")
            {
                cmbLabels.Enabled = true;
                if (blFormHasOpened)
                {
                    cmbLabels.Text = (string)cmbLabels.Items[myConfig.GetDefaultOverwriteLabelsOption() - 1];
                }
            }
            else
            {
                cmbLabels.Enabled = false;
                cmbLabels.Text = "";
            }
        }



    }
}
