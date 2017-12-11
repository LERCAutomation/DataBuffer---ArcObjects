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
                    layerPath = xmlDataExtract["LayerFolder"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'LayerFolder' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    try
                    {
                        thisLayer.Columns = aNode["Columns"].InnerText;
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
                        if (ColumnTypes.Contains(strRawText))
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
