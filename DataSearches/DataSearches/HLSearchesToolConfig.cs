﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;

using DataSearches.Properties;
using HLFileFunctions;
using HLStringFunctions;

namespace HLSearchesToolConfig
{
    class SearchesToolConfig
    {
        // Declare all the variables
        string Database;
        string LayerDir;
        string BufferLayer;
        //string EnquiriesDir;
        string RefColumn;
        string SiteColumn;
        string ReplaceChar;
        string SaveRootDir;
        string SaveFolder;
        string GISFolder;
        string LogFileName;
        bool DefaultClearLogFile;
        int DefaultBuffer;
        int DefaultUnit;
        List<string> BufferUnitOptionsDisplay = new List<string>();
        List<string> BufferUnitOptionsProcess = new List<string>();
        List<string> BufferUnitOptionsShort = new List<string>();
        string SearchLayer;
        List<string> SearchLayerExtensions = new List<string>();
        string SearchColumn;
        string AggregateColumns;
        List<string> AddSelectedLayersOptions = new List<string>();
        int DefaultAddSelectedLayers;
        List<string> OverwriteLabelOptions = new List<string>();
        int DefaultOverwriteLabels;
        List<string> CombinedSitesTableOptions = new List<string>();
        int DefaultCombinedSitesTable; // -1, 0, 1, 2 (not filled in, none, append, overwrite)
        //string CombinedSitesTableName;
        string CombinedSitesTableColumns;
        string CombinedSitesTableSuffix;
        string CombinedSitesTableFormat;

        List<string> MapLayers = new List<string>();
        List<string> MapNames = new List<string>();
        List<string> MapPrefixes = new List<string>();
        List<string> MapSuffixes = new List<string>();
        List<string> MapColumns = new List<string>();
        List<string> MapCriteria = new List<string>();
        List<string> MapGroupColumns = new List<string>();
        List<string> MapStatsColumns = new List<string>();
        List<string> MapOrderColumns = new List<string>();
        List<bool> MapIncludeDistances = new List<bool>();
        List<bool> MapIncludeRadii = new List<bool>();
        List<string> MapKeyColumns = new List<string>();
        List<string> MapFormats = new List<string>();
        List<bool> MapKeeps = new List<bool>();
        List<bool> MapLoadWarnings = new List<bool>();
        List<bool> MapPreselectLayers = new List<bool>();
        List<string> LayerFiles = new List<string>();
        List<bool> MapOverwriteLabels = new List<bool>();
        List<string> MapLabelColumns = new List<string>();
        List<string> MapLabelClauses = new List<string>();
        List<bool> MapLabelResets = new List<bool>();
        List<string> MapCombinedSiteColumns = new List<string>();
        //List<string> MapCombinedSiteCriteria = new List<string>();
        List<string> MapCombinedSiteGroupColumns = new List<string>();
        List<string> MapCombinedSiteStatsColumns = new List<string>();
        List<string> MapCombinedSiteOrderColumns = new List<string>();

        bool FoundXML;
        bool LoadedXML;
            
        // Initialise component - read XML
        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        XmlElement xmlDataSearch;
        public SearchesToolConfig()
        {
            // Open xml
            myFileFuncs = new FileFunctions();
            myStringFuncs = new StringFunctions();
            string strXMLFile = null;
            BufferUnitOptionsDisplay = new List<string>();
            BufferUnitOptionsProcess = new List<string>();
            FoundXML = false;
            LoadedXML = true;
            try
            {
                // Get the XML file
                strXMLFile = Settings.Default.XMLFile;

                // If the XML file path is blank or doesn't exist
                if (String.IsNullOrEmpty(strXMLFile) || (!myFileFuncs.FileExists(strXMLFile)))
                {
                    // Prompt the user for the correct file path
                    string strFolder = GetConfigFilePath();
                    if (!String.IsNullOrEmpty(strFolder))
                        strXMLFile = strFolder + @"\DataSearches.xml";
                }

                // Check the xml file path exists
                if (myFileFuncs.FileExists(strXMLFile))
                {
                    Settings.Default.XMLFile = strXMLFile;
                    Settings.Default.Save();
                    FoundXML = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error " + ex.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Now get all the config variables.
            // Read the file.
            if (FoundXML)
            {
                XmlDocument xmlConfig = new XmlDocument();
                try
                {
                    xmlConfig.Load(strXMLFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error in XML file; cannot load. System error message: " + ex.Message, "XML Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                string strRawText;
                XmlNode currNode = xmlConfig.DocumentElement.FirstChild; // This gets us the DataSelector.
                xmlDataSearch = (XmlElement)currNode;

                // Get all of the detail into the object

                try
                {
                    Database = xmlDataSearch["Database"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'Database' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    LayerDir = xmlDataSearch["LayerFolder"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'LayerFolder' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    BufferLayer = xmlDataSearch["BufferLayerName"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'BufferLayerName' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                
                try
                {
                    RefColumn = xmlDataSearch["RefColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'RefColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    SiteColumn = xmlDataSearch["SiteColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SiteColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    ReplaceChar = xmlDataSearch["RepChar"].InnerText;
                    if (ReplaceChar == "") ReplaceChar = " ";
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'RepChar' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    SaveRootDir = xmlDataSearch["SaveRootDir"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SaveRootDir' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    SaveFolder = xmlDataSearch["SaveFolder"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SaveFolder' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    GISFolder = xmlDataSearch["GISFolder"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'GISFolder' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    LogFileName = xmlDataSearch["LogFileName"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'LogFileName' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultClearLogFile = false;
                    string strDefaultClearLogFile = xmlDataSearch["DefaultClearLogFile"].InnerText;
                    if (strDefaultClearLogFile == "Yes")
                        DefaultClearLogFile = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultClearLogFile' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["DefaultBufferSize"].InnerText;
                    double i;
                    bool blResult = Double.TryParse(strRawText, out i);
                    if (blResult)
                        DefaultBuffer = (int)i;
                    else
                    {
                        MessageBox.Show("The entry for 'DefaultBufferSize' in the XML document is not a number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    MessageBox.Show("Could not locate the item 'DefaultBufferSize' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["DefaultBufferUnit"].InnerText;
                    int i;
                    bool blResult = Int32.TryParse(strRawText, out i);
                    if (blResult)
                        DefaultUnit = (int)i;
                    else
                    {
                        MessageBox.Show("The entry for 'DefaultBufferUnit' in the XML document is not an integer number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultBufferUnit' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["BufferUnitOptions"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'BufferUnitOptions' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                try
                {
                    char[] chrSplit1Chars = { '$' };
                    string[] liRawList = strRawText.Split(chrSplit1Chars);

                    char[] chrSplit2Chars = { ';' };
                    foreach (string strEntry in liRawList)
                    {
                        string[] strSplitEntry = strEntry.Split(chrSplit2Chars);
                        BufferUnitOptionsDisplay.Add(strSplitEntry[0]);
                        BufferUnitOptionsProcess.Add(strSplitEntry[1]);
                        BufferUnitOptionsShort.Add(strSplitEntry[2]);
                    }
                }
                catch
                {
                    MessageBox.Show("Error parsing 'BufferUnitOptions' string. Please check for correct string formatting and placement of delimiters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                

                try
                {
                    SearchLayer = xmlDataSearch["SearchLayer"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SearchLayer' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["SearchLayerExtensions"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SearchLayerExtensions' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                try
                {
                    char[] chrSplit1Chars = { ';' };
                    string[] liRawList = strRawText.Split(chrSplit1Chars);
                    foreach (string strEntry in liRawList)
                    {
                        SearchLayerExtensions.Add(strEntry);
                    }
                }
                catch
                {
                    MessageBox.Show("Error parsing 'SearchLayerExtensions' string. Please check for correct string formatting and placement of delimiters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    SearchColumn = xmlDataSearch["SearchColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SearchColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    AggregateColumns = xmlDataSearch["AggregateColumns"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'AggregateColumns' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["AddSelectedLayersOptions"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'AddSelectedLayersOptions' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                try
                {
                    char[] chrSplitChars = {';'};
                    AddSelectedLayersOptions = strRawText.Split(chrSplitChars).ToList();
                }
                catch
                {
                    MessageBox.Show("Error parsing 'AddSelectedLayersOptions' string. Please check for correct string formatting and placement of delimiters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["DefaultAddSelectedLayers"].InnerText;
                    int i;
                    bool blResult = Int32.TryParse(strRawText, out i);
                    if (blResult)
                        DefaultAddSelectedLayers = (int)i;
                    else if (strRawText == "")
                        DefaultAddSelectedLayers = -1;
                    else
                    {
                        MessageBox.Show("The entry for 'DefaultAddSelectedLayers' in the XML document is not an integer number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultAddSelectedLayers' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["OverwriteLabelOptions"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'OverwriteLabelOptions' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                try
                {
                    char[] chrSplitChars = { ';' };
                    OverwriteLabelOptions = strRawText.Split(chrSplitChars).ToList();
                }
                catch
                {
                    MessageBox.Show("Error parsing 'OverwriteLabelsOptions' string. Please check for correct string formatting and placement of delimiters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["DefaultOverwriteLabels"].InnerText;
                    int i;
                    bool blResult = Int32.TryParse(strRawText, out i);
                    if (blResult)
                        DefaultOverwriteLabels = (int)i;
                    else if (strRawText == "")
                        DefaultOverwriteLabels = -1;
                    else
                    {
                        MessageBox.Show("The entry for 'DefaultOverwriteLabels' in the XML document is not an integer number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultOverwriteLabels' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataSearch["CombinedSitesTableOptions"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'CombinedSitesTableOptions' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                try
                {
                    char[] chrSplitChars = { ';' };
                    CombinedSitesTableOptions = strRawText.Split(chrSplitChars).ToList();
                }
                catch
                {
                    MessageBox.Show("Error parsing 'CombinedSitesTableOptions' string. Please check for correct string formatting and placement of delimiters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultCombinedSitesTable = -1;
                    strRawText = xmlDataSearch["DefaultCombinedSitesTable"].InnerText;
                    int i;
                    bool blResult = Int32.TryParse(strRawText, out i);
                    if (blResult)
                        DefaultCombinedSitesTable = (int)i;
                    else if (strRawText == "")
                        DefaultCombinedSitesTable = -1;
                    else
                    {
                        MessageBox.Show("The entry for 'DefaultCombinedSitesTable' in the XML document is not an integer number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultCombinedSitesTable' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }


                try
                {
                    CombinedSitesTableColumns = xmlDataSearch["CombinedSitesTable"]["Columns"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'Columns' for entry 'CombinedSitesTable' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    CombinedSitesTableSuffix = xmlDataSearch["CombinedSitesTable"]["Suffix"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'Suffix' for entry 'CombinedSitesTable' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    CombinedSitesTableFormat = xmlDataSearch["CombinedSitesTable"]["Format"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'Format' for entry 'CombinedSitesTable' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                // All layer information.
                XmlElement MapLayerCollection = null;
                try
                {
                    MapLayerCollection = xmlDataSearch["MapLayers"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'MapLayers' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                foreach (XmlNode aNode in MapLayerCollection)
                {

                    string strName = aNode.Name;
                    strName = strName.Replace("_", " "); // Replace any underscores with spaces for better display.
                    MapLayers.Add(strName);
                    try
                    {
                        MapNames.Add(aNode["LayerName"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'LayerName' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapPrefixes.Add(aNode["Prefix"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Prefix' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }


                    try
                    {
                        MapSuffixes.Add(aNode["Suffix"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Suffix' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapColumns.Add(aNode["Columns"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Columns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strGroupColumns = aNode["GroupColumns"].InnerText;
                        // Replace the commas and any spaces.
                        strGroupColumns = myStringFuncs.getGroupColumnsFormatted(strGroupColumns);
                        MapGroupColumns.Add(strGroupColumns);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'GroupColumns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strStatsColumns = aNode["StatisticsColumns"].InnerText;
                        // Format the string
                        if (strStatsColumns != null)
                        {
                            strStatsColumns = myStringFuncs.getStatsColumnsFormatted(strStatsColumns);
                        }
                        MapStatsColumns.Add(strStatsColumns);
                        
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'StatisticsColumns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapOrderColumns.Add(aNode["OrderColumns"].InnerText); // May need to deal with.
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'OrderColumns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapCriteria.Add(aNode["Criteria"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Criteria' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        bool blIncludDistance = false;
                        string strIncludeDistance = aNode["IncludeDistance"].InnerText;
                        if (strIncludeDistance.ToLower() == "yes")
                            blIncludDistance = true;
                        MapIncludeDistances.Add(blIncludDistance);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'IncludeDistance' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        bool blIncludeRadius = false;
                        string strIncludeRadius = aNode["IncludeRadius"].InnerText;
                        if (strIncludeRadius.ToLower() == "yes")
                            blIncludeRadius = true;
                        MapIncludeRadii.Add(blIncludeRadius);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'IncludeRadius' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapKeyColumns.Add(aNode["KeyColumn"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'KeyColumn' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapFormats.Add(aNode["Format"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Format' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strKeepLayer = aNode["KeepLayer"].InnerText;
                        bool blKeepLayer = false;
                        if (strKeepLayer.ToLower() == "yes")
                            blKeepLayer = true;
                            
                        MapKeeps.Add(blKeepLayer);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'KeepLayer' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strLoadWarning = aNode["LoadWarning"].InnerText;
                        bool blLoadWarning = false;
                        if (strLoadWarning.ToLower() == "yes")
                            blLoadWarning = true;

                        MapLoadWarnings.Add(blLoadWarning);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'LoadWarning' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strPreselectLayer = aNode["PreselectLayer"].InnerText;
                        bool blPreselectLayer = false;
                        if (strPreselectLayer.ToLower() == "yes")
                            blPreselectLayer = true;

                        MapPreselectLayers.Add(blPreselectLayer);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'PreselectLayer' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        LayerFiles.Add(aNode["LayerFileName"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'LayerFileName' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strOverwriteLabels = aNode["OverwriteLabels"].InnerText;
                        bool blOverwriteLabels = false;
                        if (strOverwriteLabels.ToLower() == "yes")
                            blOverwriteLabels = true;
                        MapOverwriteLabels.Add(blOverwriteLabels);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'OverwriteLabels' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapLabelColumns.Add(aNode["LabelColumn"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'LabelColumn' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapLabelClauses.Add(aNode["LabelClause"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'LabelClause' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strMapLabelReset = aNode["LabelReset"].InnerText;
                        bool blLabelReset = false;
                        if (strMapLabelReset.ToLower() == "yes")
                            blLabelReset = true;
                        MapLabelResets.Add(blLabelReset);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'LabelReset' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapCombinedSiteColumns.Add(aNode["CombinedSitesColumns"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'CombinedSitesColumns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strGroupColumns = aNode["CombinedSitesGroupColumns"].InnerText;
                        // Replace delimiters
                        strGroupColumns = myStringFuncs.getGroupColumnsFormatted(strGroupColumns);
                        MapCombinedSiteGroupColumns.Add(strGroupColumns);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'CombinedSitesGroupColumns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        string strStatsColumns = aNode["CombinedSitesStatisticsColumns"].InnerText;
                        // Format the string
                        strStatsColumns = myStringFuncs.getStatsColumnsFormatted(strStatsColumns);
                        MapCombinedSiteStatsColumns.Add(strStatsColumns);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'CombinedSitesStatisticsColumns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        MapCombinedSiteOrderColumns.Add(aNode["CombinedSitesOrderByColumns"].InnerText); // May need to deal.
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'CombinedSitesOrderByColumns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    //try
                    //{
                    //    MapCombinedSiteCriteria.Add(aNode["CombinedSitesCriteria"].InnerText);
                    //}
                    //catch
                    //{
                    //    MessageBox.Show("Could not locate the item 'CombinedSitesCriteria' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    LoadedXML = false;
                    //    return;
                    //}
                }

            }
        }

        private string GetConfigFilePath()
        {
            // Create folder dialog.
            FolderBrowserDialog xmlFolder = new FolderBrowserDialog();

            // Set the folder dialog title.
            xmlFolder.Description = "Select folder containing 'DataSearches.xml' file ...";
            xmlFolder.ShowNewFolderButton = false;

            // Show folder dialog.
            if (xmlFolder.ShowDialog() == DialogResult.OK)
            {
                // Return the selected path.
                return xmlFolder.SelectedPath;
            }
            else
                return null;
        }

        // Return functions.
        // 1. Success criteria.
        public bool GetFoundXML()
        {
            return FoundXML;
        }

        public bool GetLoadedXML()
        {
            return LoadedXML;
        }

        // 2. General parameters.
        public string GetDatabase()
        {
            return Database;
        }

        public string GetLayerDir()
        {
            return LayerDir;
        }

        public string GetBufferLayer()
        {
            return BufferLayer;
        }

        //public string GetEnquiriesDir()
        //{
        //    return EnquiriesDir;
        //}

        public string GetRefColumn()
        {
            return RefColumn;
        }

        public string GetSiteColumn()
        {
            return SiteColumn;
        }

        public string GetReplaceChar()
        {
            return ReplaceChar;
        }

        public string GetSaveRootDir()
        {
            return SaveRootDir;
        }

        public string GetSaveFolder()
        {
            return SaveFolder;
        }

        public string GetGISFolder()
        {
            return GISFolder;
        }

        public string GetLogFileName()
        {
            return LogFileName;
        }

        public bool GetDefaultClearLogFile()
        {
            return DefaultClearLogFile;
        }

        public double GetDefaultBufferSize()
        {
            return DefaultBuffer;
        }

        public int GetDefaultBufferUnit()
        {
            return DefaultUnit;
        }

        public List<string> GetBufferUnitOptionsDisplay()
        {
            return BufferUnitOptionsDisplay;
        }

        public List<string> GetBufferUnitOptionsProcess()
        {
            return BufferUnitOptionsProcess;
        }

        public List<string> GetBufferUnitOptionsShort()
        {
            return BufferUnitOptionsShort;
        }

        public string GetSearchLayer()
        {
            return SearchLayer;
        }

        public List<string> GetSearchLayerExtensions()
        {
            return SearchLayerExtensions;
        }

        public string GetSearchColumn()
        {
            return SearchColumn;
        }

        public string GetAggregateColumns()
        {
            return AggregateColumns;
        }

        public List<string> GetAddSelectedLayerOptions()
        {
            return AddSelectedLayersOptions;
        }

        public int GetDefaultAddSelectedLayerOption()
        {
            return DefaultAddSelectedLayers;
        }

        public List<string> GetOverwriteLabelOptions()
        {
            return OverwriteLabelOptions;
        }

        public int GetDefaultOverwriteLabelsOption()
        {
            return DefaultOverwriteLabels;
        }

        public List<string> GetCombinedSitesTableOptions()
        {
            return CombinedSitesTableOptions;
        }

        public int GetDefaultCombinedSitesTable()
        {
            return DefaultCombinedSitesTable;
        }

        // 3. Setup of the combined sites table.
        //public string GetCombinedSitesTableName()
        //{
        //    return CombinedSitesTableName;
        //}

        public string GetCombinedSitesTableColumns()
        {
            return CombinedSitesTableColumns;
        }

        public string GetCombinedSitesTableSuffix()
        {
            return CombinedSitesTableSuffix;
        }

        public string GetCombinedSitesTableFormat()
        {
            return CombinedSitesTableFormat;
        }


        // 4. Setup of the map layers.
        public List<string> GetMapLayers()
        {
            return MapLayers;
        }

        public List<string> GetMapNames()
        {
            return MapNames;
        }

        public List<string> GetMapPrefixes()
        {
            return MapPrefixes;
        }

        public List<string> GetMapSuffixes()
        {
            return MapSuffixes;
        }

        public List<string> GetMapColumns()
        {
            return MapColumns;
        }

        public List<string> GetMapGroupByColumns()
        {
            return MapGroupColumns;
        }

        public List<string> GetMapStatisticsColumns()
        {
            return MapStatsColumns;
        }

        public List<string> GetMapOrderByColumns()
        {
            return MapOrderColumns;
        }

        public List<string> GetMapCriteria()
        {
            return MapCriteria;
        }

        public List<bool> getMapIncludeDistances()
        {
            return MapIncludeDistances;
        }

        public List<bool> getMapIncludeRadii()
        {
            return MapIncludeRadii;
        }

        public List<string> GetMapKeyColumns()
        {
            return MapKeyColumns;
        }

        public List<string> GetMapFormats()
        {
            return MapFormats;
        }

        public List<bool> GetMapKeepLayers()
        {
            return MapKeeps;
        }

        public List<bool> GetMapLoadWarnings()
        {
            return MapLoadWarnings;
        }

        public List<bool> GetMapPreselectLayers()
        {
            return MapPreselectLayers;
        }

        public List<string> GetMapLayerFiles()
        {
            return LayerFiles;
        }

        public List<bool> GetMapOverwriteLabels()
        {
            return MapOverwriteLabels;
        }

        public List<string> GetMapLabelColumns()
        {
            return MapLabelColumns;
        }

        public List<string> GetMapLabelClauses()
        {
            return MapLabelClauses;
        }

        public List<bool> GetMapLabelResets()
        {
            return MapLabelResets;
        }

        public List<string> GetMapCombinedSitesColumns()
        {
            return MapCombinedSiteColumns;
        }

        public List<string> GetMapCombinedSitesGroupByColumns()
        {
            return MapCombinedSiteGroupColumns;
        }

        public List<string> GetMapCombinedSitesStatsColumns()
        {
            return MapCombinedSiteStatsColumns;
        }

        public List<string> GetMapCombinedSitesOrderByColumns()
        {
            return MapCombinedSiteOrderColumns;
        }

        //public List<string> GetMapCombinedSitesCriteria()
        //{
        //    return MapCombinedSiteCriteria;
        //}
    }
}