using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using DataSearches.Properties;

using HLFileFunctions;
using HLStringFunctions;
using System.Xml;

// This configuration file reader defines how the tool behaves at start up:
// Does it show a dropdown list to choose a configuration file, or does it launch
// a default profile straight away?

namespace HLSearchesToolLaunchConfig
{
    class SearchesToolLaunchConfig
    {
        bool blChooseConfig;
        string strDefaultXML;

        bool FoundXML;
        bool LoadedXML;

        // Initialise component - read XML
        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        XmlElement xmlDataSearch;

        public SearchesToolLaunchConfig()
        {
            myFileFuncs = new FileFunctions();
            myStringFuncs = new StringFunctions();
            string strXMLFile = null;
            FoundXML = false;
            LoadedXML = true;

            // Open and read the XML file. 
            try
            {
                // Get the XML file
                strXMLFile = Settings.Default.XMLFile;

                // If the XML file path is blank or doesn't exist
                if (String.IsNullOrEmpty(strXMLFile) || (!myFileFuncs.FileExists(strXMLFile)))
                {
                    // Prompt the user for the correct file path. File name is always the same.
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

            // Now get the variables.
            if (FoundXML)
            {
                // Load the XML file into memory.
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
                XmlNode currNode = xmlConfig.DocumentElement.FirstChild; // This gets us the InitialConfig.
                xmlDataSearch = (XmlElement)currNode;

                // Get the multiple choice variable.
                try
                {
                    blChooseConfig = false;
                    strRawText = xmlDataSearch["ChooseXML"].InnerText;
                    if (strRawText.ToLower() == "yes" || strRawText.ToLower() == "y")
                    {
                        blChooseConfig = true;
                    }
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ChooseXML' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                // Get the default XML file.
                try
                {
                    strDefaultXML = xmlDataSearch["DefaultProfile"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultProfile' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
            }
        }


        // Return variables.
        public bool XMLFound
        {
            get
            {
                return FoundXML;
            }
        }

        public bool XMLLoaded
        {
            get
            {
                return LoadedXML;
            }
        }

        public bool ChooseConfig
        {
            get
            {
                return blChooseConfig;
            }
        }

        public string DefaultXML
        {
            get
            {
                return strDefaultXML;
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
    }

    
}
