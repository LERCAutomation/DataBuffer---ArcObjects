using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using DataBuffer;
using DataBuffer.Properties;
using HLFileFunctions;
using HLStringFunctions;


namespace HLDataBufferConfig
{
    class DataBufferConfig
    {
        // Declare all the variables.
        // Environment and menu variables.
        private string logFilePath;
        private bool defaultClearLog;
        private string defaultPath;
        private string layerPath;
        private string tempFilePath;

        //private string outColumnDefs;
        public string LogFilePath 
        { 
            get
            {
                return logFilePath;
            } 
        }

        public bool DefaultClearLog 
        { 
            get
            {
                return defaultClearLog;
            } 
        }
        public string DefaultPath 
        {
            get
            {
                return defaultPath;
            }
        }

        public string LayerPath
        {
            get
            {
                return layerPath;
            }
        }

        public string TempFilePath
        {
            get
            {
                return tempFilePath;
            }
        }

        private MapLayers inputLayers = new MapLayers();
        public MapLayers InputLayers
        {
            get
            {
                return inputLayers;
            }
        }

        private OutputLayer outputLayer = new OutputLayer();
        public OutputLayer OutputLayer
        {
            get
            {
                return outputLayer;
            }
        }

        private bool foundXML;
        private bool loadedXML;
        public bool FoundXML
        {
            get
            {
                return foundXML;
            }
        }
        public bool LoadedXML 
        {
            get
            {
                return loadedXML;
            }
        }

        private FileFunctions myFileFuncs;
        private StringFunctions myStringFuncs;
        private XmlElement xmlDataExtract;

        public DataBufferConfig()
        {
            // Open XML
            myFileFuncs = new FileFunctions();
            myStringFuncs = new StringFunctions();
            string strXMLFile = null;
            foundXML = false;
            loadedXML = true;
            
            
            try
            {
                // Get the XML file
                strXMLFile = Settings.Default.XMLFile;
                if (String.IsNullOrEmpty(strXMLFile) || (!myFileFuncs.FileExists(strXMLFile))) // Can't find it or doesn't exist
                {
                    // Prompt the user for the location of the file.
                    string strFolder = GetConfigFilePath();
                    if (!String.IsNullOrEmpty(strFolder))
                        strXMLFile = strFolder + @"\DataBuffer.xml";
                }

                // check the xml file path exists
                if (myFileFuncs.FileExists(strXMLFile))
                {
                    Settings.Default.XMLFile = strXMLFile;
                    Settings.Default.Save();
                    foundXML = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error " + ex.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            // Go forth and obtain all information.
            // Firstly, read the file.
            if (foundXML)
            {
                XmlDocument xmlConfig = new XmlDocument();
                try
                {
                    xmlConfig.Load(strXMLFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error in XML file; cannot load. System error message: " + ex.Message, "XML Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }
                string strRawText;
                XmlNode currNode = xmlConfig.DocumentElement.FirstChild; // This gets us the DataSelector.
                xmlDataExtract = (XmlElement)currNode;

                // XML loaded successfully; get all of the detail in the Config object.
                
                try
                {
                    logFilePath = xmlDataExtract["LogFilePath"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'LogFilePath' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }

                try
                {
                    defaultClearLog = false;
                    string strDefaultClearLogFile = xmlDataExtract["DefaultClearLogFile"].InnerText;
                    if (strDefaultClearLogFile == "Yes")
                        defaultClearLog = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultClearLogFile' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }

                try
                {
                    defaultPath = xmlDataExtract["DefaultPath"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultPath' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }

                try
                {
                    layerPath = xmlDataExtract["LayerPath"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'LayerPath' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }

                try
                {
                    tempFilePath = xmlDataExtract["TempFilePath"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'TempFilePath' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }

                // Locate the GIS Layers.
                XmlElement MapLayerCollection = null;
                try
                {
                    MapLayerCollection = xmlDataExtract["InLayers"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'InLayers' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }
                // Now cycle through them.
                foreach (XmlNode aNode in MapLayerCollection)
                {
                    MapLayer thisLayer = new MapLayer();
                    thisLayer.DisplayName = aNode.Name;

                    try
                    {
                        thisLayer.LayerName = aNode["LayerName"].InnerText;
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'LayerName' for map layer " + thisLayer.DisplayName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    // Sort out the columns. This is pretty involved.
                    try
                    {
                        InputColumns theInputColumns = new InputColumns();
                        string strColumnList = aNode["Columns"].InnerText;
                        // We have the format (inputColumn1 "outputColumn1", inputColumn2, inputColumn3, "outputColumn3", "inputText" "outputColumn4", ...)
                        // Firstly split the list at the commas.
                        List<string> strColumnDefList = strColumnList.Split(',').ToList();
                        // Go through these and sort out what's what.
                        foreach (string aColumnDef in strColumnDefList)
                        {
                            InputColumn thisInputColumn = new InputColumn();
                            // Check if the first character is a "\"". If so, we deal with it slightly differently.
                            string strColumnDef = aColumnDef.Trim(); // Remove any spaces.
                            List<string> strColItems = new List<string>();
                            if (strColumnDef.Substring(0, 1) == "\"")
                            {
                                // find the first entry.
                                int position; // First character is a '"' so we don't want to find that.
                                int start = 0;
                                // Extract the items from the string.
                                position = strColumnDef.IndexOf('\"', start + 1);
                                if (position == 0) position = 1;
                                if (position > 0)
                                {
                                    string strResult = strColumnDef.Substring(start, position - start + 1).Trim();
                                    strColItems.Add(strResult);
                                    //start = position;
                                }
                                // The second item is split by string. 
                                List<string> strAllEntries = strColumnDef.Split(' ').ToList();
                                string theEntry = strAllEntries[strAllEntries.Count - 1]; // Last entry.
                                strColItems.Add(theEntry.Trim('"')); // Trim quotes if they are there.

                            }
                            else
                            {
                                // Split at space.
                                strColItems = strColumnDef.Split(' ').ToList();
                            }
                            // Test to see how many elements.
                            if (strColItems.Count == 1)
                            {
                                thisInputColumn.InputName = strColItems[0].Trim();
                                thisInputColumn.OutputName = strColItems[0].Trim(); // They are both the same.
                            }
                            else if (strColItems.Count == 2)
                            {
                                thisInputColumn.InputName = strColItems[0].Trim();
                                thisInputColumn.OutputName = strColItems[1].Trim('"'); // Trim quotes if they are there
                            }
                            else
                            {
                                // More than two elements; that's not right.
                                MessageBox.Show("The column entry " + strColItems[0] + " for map layer " + thisLayer.DisplayName + " in the XML file contains " + strColItems.Count.ToString() + " items. It should only have two.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                loadedXML = false;
                                return;
                            }

                            theInputColumns.Add(thisInputColumn);
                        }
                        thisLayer.InputColumns = theInputColumns;
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Columns' for map layer " + thisLayer.DisplayName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    try
                    {
                        thisLayer.WhereClause = aNode["WhereClause"].InnerText;
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'WhereClause' for map layer " + thisLayer.DisplayName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    try
                    {
                        thisLayer.SortOrder = aNode["SortOrder"].InnerText;
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'SortOrder' for map layer " + thisLayer.DisplayName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    try
                    {
                        int a;
                        bool blResult = int.TryParse(aNode["BufferSize"].InnerText, out a);
                        if (blResult)
                            thisLayer.BufferSize = a;
                        else
                        {
                            MessageBox.Show("Could not locate the item 'BufferSize' for map layer " + thisLayer.DisplayName + " in the XML file, or the item is not an integer", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            loadedXML = false;
                            return;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'BufferSize' for map layer " + thisLayer.DisplayName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    try
                    {
                        int a;
                        bool blResult = int.TryParse(aNode["DissolveSize"].InnerText, out a);
                        if (blResult)
                            thisLayer.DissolveSize = a;
                        else
                        {
                            MessageBox.Show("Could not locate the item 'DissolveSize' for map layer " + thisLayer.DisplayName + " in the XML file, or the item is not an integer", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            loadedXML = false;
                            return;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'BufferSize' for map layer " + thisLayer.DisplayName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }


                    // if everything is correct, add to the list.
                    if (loadedXML)
                        inputLayers.Add(thisLayer);
                }



                // Now get the output layer definition
                XmlElement OutLayerDef = null;
                try
                {
                    OutLayerDef = xmlDataExtract["OutLayer"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'OutLayer' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                } 

                // Get to the columns
                XmlNode ColumnNode = null;
                try
                {
                    ColumnNode = OutLayerDef["Columns"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'Columns' for the OutLayer in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }

                    
                // Get all the columns.
                OutputColumns theOutputColumns = new OutputColumns();
                foreach (XmlNode aNode in ColumnNode)
                {
                    OutputColumn thisColumn = new OutputColumn();
                    thisColumn.ColumnName = aNode.Name;

                    try
                    {
                        thisColumn.ColumnName = aNode["ColumnName"].InnerText;
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'ColumnName' for output column " + thisColumn.ColumnTag + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    // List of accepted types.
                    List<string> ColumnTypes = new List<string>() { "key", "cluster", "first", "common", "min", "max", "range" };
                    try
                    {
                        strRawText = aNode["ColumnType"].InnerText.ToLower();
                        if (ColumnTypes.Contains(strRawText))
                        {
                            thisColumn.ColumnType = strRawText; // Always lower case.

                            // Now also add this type to the relevant output column in ALL the input layers.
                            foreach (MapLayer aLayer in inputLayers)
                            {
                                // Find the output column with the same name.
                                bool blFoundIt = false;
                                foreach (InputColumn aColumn in aLayer.InputColumns)
                                {
                                    if (aColumn.OutputName == thisColumn.ColumnName)
                                    {
                                        aColumn.ColumnType = strRawText;
                                        blFoundIt = true;
                                        break;
                                    }
                                }
                                if (!blFoundIt)
                                {
                                    MessageBox.Show("The output column " + thisColumn.ColumnName + " was not found for map layer " + aLayer.LayerName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    loadedXML = false;
                                    return;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("The value for 'ColumnType' for output column " + thisColumn.ColumnTag + " in the XML file is not valid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            loadedXML = false;
                            return;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'ColumnType' for output column " + thisColumn.ColumnTag + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    List<string> FieldTypes = new List<string>() { "TEXT", "FLOAT", "DOUBLE", "SHORT", "LONG", "DATE"};
                    try
                    {
                        strRawText = aNode["FieldType"].InnerText.ToUpper();
                        if (FieldTypes.Contains(strRawText))
                        {
                            thisColumn.FieldType = strRawText; // Always upper case.
                        }
                        else
                        {
                            MessageBox.Show("The value for 'FieldType' for output column " + thisColumn.ColumnTag + " in the XML file is not valid", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            loadedXML = false;
                            return;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'ColumnType' for output column " + thisColumn.ColumnTag + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    try
                    {
                        int a;
                        bool blResult = int.TryParse(aNode["ColumnLength"].InnerText, out a);
                        if (blResult)
                            thisColumn.ColumnLength = a;
                        else
                        {
                            MessageBox.Show("Could not locate the item 'ColumnLength' for map layer " + thisColumn.ColumnTag + " in the XML file, or the item is not an integer", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            loadedXML = false;
                            return;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'ColumnLength' for map layer " + thisColumn.ColumnTag + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                        return;
                    }

                    if (loadedXML)
                    {
                        theOutputColumns.Add(thisColumn);
                    }
                }

                // Add the columns to the output layer definition
                outputLayer.OutputColumns = theOutputColumns;

                // Get the rest of the output layer definition
                try
                {
                    outputLayer.LayerFile = OutLayerDef["LayerFile"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'LayerFile' for the OutLayer in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }

                try
                {
                    List<string> OutputFormatList = new List<string>() { "shape", "gdb" };
                    strRawText = OutLayerDef["OutputFormat"].InnerText.ToLower();
                    if (OutputFormatList.Contains(strRawText))
                    {
                        outputLayer.Format = strRawText;
                    }
                    else
                    {
                        MessageBox.Show("The entry for the output layer's OutputFormat in the XML file is not valid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        loadedXML = false;
                    }
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'OutputFormat' for the OutLayer in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    loadedXML = false;
                    return;
                }
            }
        }

        


        private string GetConfigFilePath()
        {
            // Create folder dialog.
            FolderBrowserDialog xmlFolder = new FolderBrowserDialog();

            // Set the folder dialog title.
            xmlFolder.Description = "Select folder containing 'DataExtractor.xml' file ...";
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
