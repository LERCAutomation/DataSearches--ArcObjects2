﻿// DataSearches is an ArcGIS add-in used to extract biodiversity
// and conservation area information from ArcGIS based on a radius around a feature.
//
// Copyright © 2016-2017 SxBRC, 2017-2018 TVERC
//
// This file is part of DataSearches.
//
// DataSearches is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataSearches is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataSearches.  If not, see <http://www.gnu.org/licenses/>.

namespace DataSearches
{
    partial class frmDataSearches
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDataSearches));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtLocation = new System.Windows.Forms.TextBox();
            this.lstLayers = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtBufferSize = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbUnits = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.chkClearLog = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cmbAddLayers = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbLabels = new System.Windows.Forms.ComboBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cmbCombinedSites = new System.Windows.Forms.ComboBox();
            this.txtOrganisation = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Search Reference:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Site Name";
            // 
            // txtLocation
            // 
            this.txtLocation.Enabled = false;
            this.txtLocation.Location = new System.Drawing.Point(12, 71);
            this.txtLocation.Name = "txtLocation";
            this.txtLocation.Size = new System.Drawing.Size(393, 20);
            this.txtLocation.TabIndex = 3;
            // 
            // lstLayers
            // 
            this.lstLayers.FormattingEnabled = true;
            this.lstLayers.Location = new System.Drawing.Point(12, 118);
            this.lstLayers.Name = "lstLayers";
            this.lstLayers.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstLayers.Size = new System.Drawing.Size(182, 277);
            this.lstLayers.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(214, 117);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Buffer Size:";
            // 
            // txtBufferSize
            // 
            this.txtBufferSize.Location = new System.Drawing.Point(217, 136);
            this.txtBufferSize.Name = "txtBufferSize";
            this.txtBufferSize.Size = new System.Drawing.Size(188, 20);
            this.txtBufferSize.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(214, 164);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Buffer Units:";
            // 
            // cmbUnits
            // 
            this.cmbUnits.FormattingEnabled = true;
            this.cmbUnits.Location = new System.Drawing.Point(217, 180);
            this.cmbUnits.Name = "cmbUnits";
            this.cmbUnits.Size = new System.Drawing.Size(188, 21);
            this.cmbUnits.TabIndex = 8;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 102);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(90, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Layers to Search:";
            // 
            // chkClearLog
            // 
            this.chkClearLog.AutoSize = true;
            this.chkClearLog.Checked = true;
            this.chkClearLog.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkClearLog.Location = new System.Drawing.Point(12, 401);
            this.chkClearLog.Name = "chkClearLog";
            this.chkClearLog.Size = new System.Drawing.Size(96, 17);
            this.chkClearLog.TabIndex = 11;
            this.chkClearLog.Text = "Clear Log File?";
            this.chkClearLog.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(214, 221);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(144, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "Add Selected Layers to Map:";
            // 
            // cmbAddLayers
            // 
            this.cmbAddLayers.FormattingEnabled = true;
            this.cmbAddLayers.Location = new System.Drawing.Point(217, 237);
            this.cmbAddLayers.Name = "cmbAddLayers";
            this.cmbAddLayers.Size = new System.Drawing.Size(188, 21);
            this.cmbAddLayers.TabIndex = 13;
            this.cmbAddLayers.SelectedIndexChanged += new System.EventHandler(this.cmbAddLayers_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(214, 261);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(113, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Overwrite Map Labels:";
            // 
            // cmbLabels
            // 
            this.cmbLabels.Enabled = false;
            this.cmbLabels.FormattingEnabled = true;
            this.cmbLabels.Location = new System.Drawing.Point(217, 277);
            this.cmbLabels.Name = "cmbLabels";
            this.cmbLabels.Size = new System.Drawing.Size(188, 21);
            this.cmbLabels.TabIndex = 15;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(252, 395);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 17;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(333, 395);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 18;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(12, 24);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(264, 20);
            this.txtSearch.TabIndex = 0;
            this.txtSearch.TextChanged += new System.EventHandler(this.txtSearch_TextChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(214, 323);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(147, 13);
            this.label8.TabIndex = 19;
            this.label8.Text = "Create Combined Sites Table:";
            // 
            // cmbCombinedSites
            // 
            this.cmbCombinedSites.FormattingEnabled = true;
            this.cmbCombinedSites.Location = new System.Drawing.Point(214, 339);
            this.cmbCombinedSites.Name = "cmbCombinedSites";
            this.cmbCombinedSites.Size = new System.Drawing.Size(191, 21);
            this.cmbCombinedSites.TabIndex = 20;
            // 
            // txtOrganisation
            // 
            this.txtOrganisation.Enabled = false;
            this.txtOrganisation.Location = new System.Drawing.Point(301, 47);
            this.txtOrganisation.Name = "txtOrganisation";
            this.txtOrganisation.Size = new System.Drawing.Size(104, 20);
            this.txtOrganisation.TabIndex = 21;
            this.txtOrganisation.Visible = false;
            // 
            // frmDataSearches
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(420, 430);
            this.Controls.Add(this.txtOrganisation);
            this.Controls.Add(this.cmbCombinedSites);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtSearch);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cmbLabels);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cmbAddLayers);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.chkClearLog);
            this.Controls.Add(this.cmbUnits);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtBufferSize);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lstLayers);
            this.Controls.Add(this.txtLocation);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label6);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmDataSearches";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = "Data Searches " + string.Format("{0}.{1}.{2}", version.Major, version.Minor, version.Build);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLocation;
        private System.Windows.Forms.ListBox lstLayers;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtBufferSize;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbUnits;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox chkClearLog;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cmbAddLayers;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbLabels;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cmbCombinedSites;
        private System.Windows.Forms.TextBox txtOrganisation;
    }
}