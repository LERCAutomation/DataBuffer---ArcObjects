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
using System.Diagnostics; // Allows Process to be called (for Notepad)
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using HLBufferToolLaunchConfig;
using HLDataBufferConfig;
using HLArcMapModule;
using HLFileFunctions;
using DataBuffer.Properties;
using HLStringFunctions;

using ESRI.ArcGIS.Framework;

namespace DataBuffer
{
    public partial class frmDataBuffer : Form
    {
        IApplication theApplication;
        DataBufferRoutine myDataBufferFuncs;
        ArcMapFunctions myArcMapFuncs;
        FileFunctions myFileFuncs;
        DataBufferConfig myConfig;
        BufferToolLaunchConfig myLaunchConfig;
        StringFunctions myStringFuncs;
        string strConfigFile = "";

        bool blOpenForm; // this tracks all the way through initialisation whether the form should open.

        public frmDataBuffer()
        {
            blOpenForm = true;
            InitializeComponent();

            myLaunchConfig = new BufferToolLaunchConfig();
            myFileFuncs = new FileFunctions();
            myStringFuncs = new StringFunctions();
            if (!myLaunchConfig.XMLFound)
            {
                MessageBox.Show("XML file 'DataSelector.xml' not found; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }
            if (!myLaunchConfig.XMLLoaded)
            {
                MessageBox.Show("Error loading XML File 'DataSelector.xml'; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            if (blOpenForm)
            {
                string strXMLFolder = myFileFuncs.GetDirectoryName(Settings.Default.XMLFile);
                bool blOnlyDefault = true;
                int intCount = 0;
                if (myLaunchConfig.ChooseConfig) // If we are allowed to choose, check if there are multiple profiles. 
                // If there is only the default XML file in the directory, launch the form. Otherwise the user has to choose.
                {
                    foreach (string strFileName in myFileFuncs.GetAllFilesInDirectory(strXMLFolder))
                    {
                        if (myFileFuncs.GetFileName(strFileName).ToLower() != "dataselector.xml" && myFileFuncs.GetExtension(strFileName).ToLower() == "xml")
                        {
                            // is it the default?
                            intCount++;
                            if (myFileFuncs.GetFileName(strFileName) != myLaunchConfig.DefaultXML)
                            {
                                blOnlyDefault = false;
                            }
                        }
                    }
                    if (intCount > 1)
                    {
                        blOnlyDefault = false;
                    }
                }
                if (myLaunchConfig.ChooseConfig && !blOnlyDefault)
                {
                    // User has to choose the configuration file first.

                    using (var myConfigForm = new frmChooseConfig(strXMLFolder, myLaunchConfig.DefaultXML))
                    {
                        var result = myConfigForm.ShowDialog();
                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            strConfigFile = strXMLFolder + "\\" + myConfigForm.ChosenXMLFile;
                        }
                        else
                        {
                            MessageBox.Show("No XML file was chosen; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            blOpenForm = false;
                        }
                    }

                }
                else
                {
                    strConfigFile = strXMLFolder + "\\" + myLaunchConfig.DefaultXML; // don't allow the user to choose, just use the default.
                    // Just check it exists, though.
                    if (!myFileFuncs.FileExists(strConfigFile))
                    {
                        MessageBox.Show("The default XML file '" + myLaunchConfig.DefaultXML + "' was not found in the XML directory; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        blOpenForm = false;
                    }
                }
            }

            if (blOpenForm)
            {
                // Firstly let's read the XML.
                myConfig = new DataBufferConfig(strConfigFile); // Must now pass the correct XML name.

                // Did we find the XML?
                if (!myConfig.FoundXML)
                {
                    MessageBox.Show("XML file not found; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }

                // Did it load OK?
                else if (!myConfig.LoadedXML)
                {
                    MessageBox.Show("Error loading XML File; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }
            }

            // Close the form if there are any errors at this point.
            if (!blOpenForm)
            {
                Load += (s, e) => Close();
                return;
            }

            // We're all set to show the form. Set it up.
            // Initialise all the helper classes.
            theApplication = ArcMap.Application;
            myArcMapFuncs = new ArcMapFunctions(theApplication);
            myDataBufferFuncs = new DataBufferRoutine(theApplication, strConfigFile);
            myFileFuncs = new FileFunctions();

            // Now fill up the menu with the required layers.
            // Firstly check for missing layers.
            MapLayers theInputLayers = myConfig.InputLayers;

            List<string> MissingLayerList = new List<string>();
            foreach (MapLayer aLayer in theInputLayers)
            {
                if (!myArcMapFuncs.LayerExists(aLayer.LayerName)) // We do not accept group layers.
                    MissingLayerList.Add(aLayer.LayerName);
                else
                    lstInput.Items.Add(aLayer.DisplayName);
            }

            // Tell the user that there's a problem if there is one.
            if (MissingLayerList.Count > 0)
            {
                string strMessage = "Warning: ";
                if (MissingLayerList.Count == 1)
                    strMessage = "the layer " + MissingLayerList[0] + " is not loaded in the Table of Contents.";
                else if (MissingLayerList.Count > 1)
                {
                    strMessage = "the following layers are not loaded in the Table of Contents: ";
                    foreach (string aLayer in MissingLayerList)
                    {
                        strMessage = strMessage + aLayer + ", ";
                    }
                    strMessage = strMessage.Substring(0, strMessage.Length - 2) + "."; // Trim the last comma and space; add a full stop.
                }
                MessageBox.Show(strMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // Set the default for clear log file.
            chkClearLog.Checked = myConfig.DefaultClearLog;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            // Do we have a selection? If not, remind the user and exit.
            if (lstInput.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one layer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = Cursors.Default;
                return;
            }

            // Fix any illegal characters in the user name string
            string strUserID = myStringFuncs.StripIllegals(Environment.UserName, "_", false);

            // Define the log file
            string strLogFile = myConfig.LogFilePath + "\\DataBuffer_" + strUserID + ".log";

            // Delete if requested
            if (chkClearLog.Checked)
            {
                bool blDeleted = myFileFuncs.DeleteFile(strLogFile);
                if (!blDeleted)
                {
                    MessageBox.Show("Cannot delete log file. Please make sure it is not open in another window");
                    return;
                }
            }

            //this.btnOK.Enabled = false;
            this.Enabled = false;
            myArcMapFuncs.ToggleDrawing(false);
            myArcMapFuncs.ToggleTOC(false);

            // Request the output file from the user.
            string strOutputFile = "None";
            bool blDone = false;
            while (!blDone)
            {
                strOutputFile = myArcMapFuncs.GetOutputFileName(myConfig.OutputLayer.Format, myConfig.DefaultPath);
                if (strOutputFile != "None")
                {
                    // Does this output file already exist?
                    if (myArcMapFuncs.FeatureclassExists(strOutputFile))
                    {
                        DialogResult confirmOverwrite = MessageBox.Show("The output file already exists. Do you want to overwrite it?", "Data Buffer", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (confirmOverwrite == System.Windows.Forms.DialogResult.Yes)
                        {
                            // Delete the existing file
                            //myArcMapFuncs.DeleteFeatureclass(strOutputFile);
                            blDone = true;
                        }
                    }
                    else
                    {
                        blDone = true; // It's a new file
                    }
                }
                else
                {
                    // User pressed Cancel - return to main menu.
                    this.Cursor = Cursors.Default;
                    this.Enabled = true;
                    myArcMapFuncs.ToggleDrawing(true);
                    myArcMapFuncs.ToggleTOC(true);
                    return;
                }
            }



            // Find the selected map layers.
            MapLayers Selectedlayers = new MapLayers();
            MapLayers AllMapLayers = myConfig.InputLayers;
            foreach (string strSelectedItem in lstInput.SelectedItems)
            {
                // Find the relevant map layer and add it to the collection.
                foreach (MapLayer aLayer in AllMapLayers)
                {
                    if (aLayer.DisplayName == strSelectedItem)
                        Selectedlayers.Add(aLayer);
                }
            }


            // Set up the data buffer functions class.
            myDataBufferFuncs = new DataBufferRoutine(theApplication, strConfigFile);
            myDataBufferFuncs.InputLayers = Selectedlayers;
            myDataBufferFuncs.OutputLayer = myConfig.OutputLayer;

            // *** RUN THE FUNCTION ***
            long intResult = myDataBufferFuncs.RunDataBuffer(strOutputFile, strLogFile, this);

            // Report the result
            if (intResult != -999)
            {
                myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");
                myFileFuncs.WriteLine(strLogFile, "Process complete. The new buffer layer has " + intResult.ToString() + " records.");
                myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");
            }
            else
            {
                myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");
                myFileFuncs.WriteLine(strLogFile, "Process could not complete successfully due to errors.");
                myFileFuncs.WriteLine(strLogFile, "---------------------------------------------------------------------------");
            }

            // Finish off.
            this.Cursor = Cursors.Default;
            this.Enabled = true;
            if (intResult != -999)
            {
                DialogResult dlResult = MessageBox.Show("Process complete. Do you wish to close the form?", "Data Buffer", MessageBoxButtons.YesNo);
                if (dlResult == System.Windows.Forms.DialogResult.Yes)
                    this.Close();
                else
                    this.BringToFront();
            }
            else
            {
                this.BringToFront();
            }
            Process.Start("notepad.exe", strLogFile);

            // Any required tidying up.
            Selectedlayers = null;
            AllMapLayers = null;
            myArcMapFuncs.ToggleDrawing(true);
            myArcMapFuncs.ToggleTOC(true);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void UpdateStatus(string aNewStatusBottom, string aNewStatusTop = "")
        {
            if (aNewStatusTop != "")
                slStatus1.Text = aNewStatusTop;
            if (aNewStatusBottom == ".")
                slStatus2.Text = slStatus2.Text + ".";
            else
                slStatus2.Text = aNewStatusBottom;
            this.Update();
        }
    }
}
