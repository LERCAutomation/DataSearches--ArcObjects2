using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using HLFileFunctions;


namespace DataSearches
{
    public partial class frmChooseConfig : Form
    {
        public string ChosenXMLFile { get; set; }
        FileFunctions myFileFuncs = new FileFunctions();
        public frmChooseConfig(string XMLFolder, string DefaultXMLName)
        {
            InitializeComponent();
            // Get the files in the XML directory.
            List<string> myFileList = myFileFuncs.GetAllFilesInDirectory(XMLFolder);
            List<string> myFilteredFiles = new List<string>();
            bool blDefaultFound = false;
            foreach (string aFile in myFileList)
            {
                // Add it if it's not DataSearches.xml.
                string aFileName = myFileFuncs.GetFileName(aFile);
                if (aFileName.ToLower() != "datasearches.xml" && myFileFuncs.GetExtension(aFile).ToLower() == "xml")
                {
                    myFilteredFiles.Add(aFileName);
                    if (aFileName.ToLower() == DefaultXMLName.ToLower())
                        blDefaultFound = true;
                }
            }
            myFileList.Sort();
            // Add the files to the dropdown list.
            foreach (string aFile in myFilteredFiles)
            {
                cmbChooseXML.Items.Add(aFile);
            }
            // Now select the default.
            if (blDefaultFound)
                cmbChooseXML.SelectedItem = DefaultXMLName;
        }

        private void frmChooseConfig_Load(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.ChosenXMLFile = cmbChooseXML.Text;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
