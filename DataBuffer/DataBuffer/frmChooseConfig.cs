// DataBuffer is an ArcGIS add-in used to create 'species alert'
// layers from existing species data.
//
// Copyright © 2017 SxBRC, 2017-2018 TVERC
//
// This file is part of DataBuffer.
//
// DataBuffer is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataBuffer is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataBuffer.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using HLFileFunctions;


namespace DataBuffer
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
                // Add it if it's not DataBuffer.xml.
                string aFileName = myFileFuncs.GetFileName(aFile);
                if (aFileName.ToLower() != "databuffer.xml" && myFileFuncs.GetExtension(aFile).ToLower() == "xml")
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
                cmbChooseXML.Items.Add(myFileFuncs.ReturnWithoutExtension(aFile));
            }
            // Now select the default.
            if (blDefaultFound)
                cmbChooseXML.SelectedItem = myFileFuncs.ReturnWithoutExtension(DefaultXMLName);
        }

        private void frmChooseConfig_Load(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.ChosenXMLFile = cmbChooseXML.Text + ".xml";
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
