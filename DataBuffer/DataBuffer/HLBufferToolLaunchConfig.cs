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
using System.Linq;
using System.Text;

using System.Windows.Forms;

using DataBuffer.Properties;

using HLFileFunctions;
using HLStringFunctions;
using System.Xml;

// This configuration file reader defines how the tool behaves at start up:
// Does it show a dropdown list to choose a configuration file, or does it launch
// a default profile straight away?

namespace HLBufferToolLaunchConfig
{
    class BufferToolLaunchConfig
    {
        bool blChooseConfig;
        string strDefaultXML = "DefaultProfile.xml";

        bool FoundXML;
        bool LoadedXML;

        // Initialise component - read XML
        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        XmlElement xmlDataBuffer;

        public BufferToolLaunchConfig()
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
                        strXMLFile = strFolder + @"\DataBuffer.xml";
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
                xmlDataBuffer = (XmlElement)currNode;

                // Get the multiple choice variable.
                try
                {
                    blChooseConfig = false;
                    strRawText = xmlDataBuffer["ChooseXML"].InnerText;
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
                    strRawText = xmlDataBuffer["DefaultProfile"].InnerText;
                    if (strRawText != "")
                        strDefaultXML = strRawText; // If there is an entry; otherwise use the default.
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
            xmlFolder.Description = "Select folder containing 'DataBuffer.xml' file ...";
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
