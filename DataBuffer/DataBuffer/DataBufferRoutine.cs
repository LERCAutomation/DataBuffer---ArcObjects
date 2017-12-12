using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;

using HLArcMapModule;
using HLDataBufferConfig;
using HLFileFunctions;
using HLStringFunctions;

using ESRI.ArcGIS.Framework;

namespace DataBuffer
{
    class DataBufferRoutine
    {
        // Properties
        public MapLayers InputLayers { get; set; }
        public OutputLayer OutputLayer { get; set; }
        public string OutputFile { get; set; }

        private ArcMapFunctions myArcMapFuncs;
        private DataBufferConfig myConfig;
        private FileFunctions myFileFuncs;

        public DataBufferRoutine(IApplication theApplication) // comes from the interface hence through the form.
        {
            // Get all the essentials up and running
            myArcMapFuncs = new ArcMapFunctions(theApplication);
            myConfig = new DataBufferConfig();
            myFileFuncs = new FileFunctions();
        }

        public long RunDataBuffer(string anOutputFile, string aLogFile)
        {
            // This uses the settings XML to retrieve the layerfolder MAYBE
            // so they do not need to be passed as a setting.
            // Inputlayers, outputlayers and output file should already have been set.

            // Note this has been set up to be rather test-driven. While it is very hard to formally
            // unit test ArcGIS add-ins, this code attempts to be as airtight in terms of catching
            // all eventualities as I could get it.

            long lngResult = -999; // Initially assume failure.
            bool blTest = false; // Testing variable.

            // Check input. None of these should ever be fired.
            if (InputLayers == null)
            {
                MessageBox.Show("The input layers have not been defined.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lngResult;
            }
            else if ( OutputLayer == null)
            {
                MessageBox.Show("The output layer has not been defined.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lngResult;
            }
            else if (OutputFile == "")
            {
                MessageBox.Show("The output file has not been defined.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lngResult;
            }

            // All good to go.
            // 1. Start the log file. 
            string strLogFile = aLogFile;
            blTest = myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");
            if (!blTest) // We only test access to log file once.
            {
                myFileFuncs.WriteLine(strLogFile, "Could not write to log file " + aLogFile);
                MessageBox.Show("Unable to write to log file", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lngResult;
            }
            myFileFuncs.WriteLine(strLogFile, "Process started");
            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");

            string strMessage = "The following layers are being processed: ";
            foreach (MapLayer aLayer in InputLayers)
            {
                strMessage = strMessage + aLayer.DisplayName + ", ";
            }
            strMessage = strMessage.Substring(0, strMessage.Length - 2) + ".";
            myFileFuncs.WriteLine(strLogFile, strMessage);
            myFileFuncs.WriteLine(strLogFile, "The output layer is " + anOutputFile + ".");
            strMessage = "The following output columns are included: ";
            foreach (OutputColumn aCol in OutputLayer.OutputColumns)
            {
                strMessage = strMessage + aCol.ColumnName + ", ";
            }
            strMessage = strMessage.Substring(0, strMessage.Length - 2) + ".";
            myFileFuncs.WriteLine(strLogFile, strMessage);

            // 2. Create the empty output FC using correct definitions.
            string strOutFCName = myFileFuncs.GetFileName(anOutputFile);
            string strOutFolder = myFileFuncs.GetDirectoryName(anOutputFile);

            // 2a. Create a temporary directory, and a GDB inside.
            string strTempDir = myConfig.TempFilePath;
            
            if (!myFileFuncs.DirExists(strTempDir))
            {
                try
                {
                    DirectoryInfo tempDir = Directory.CreateDirectory(strTempDir);
                    tempDir.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }
                catch
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not create temporary folder " + strTempDir);
                    MessageBox.Show("Temporary folder could not be created", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                myFileFuncs.WriteLine(strLogFile, "Temporary folder created.");
            }
            string strTempGDB = strTempDir + @"\Temp.gdb";
            if (!myFileFuncs.DirExists(strTempGDB))
            {
                blTest = myArcMapFuncs.CreateWorkspace(strTempGDB);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not create temporary geodatabase " + strTempGDB);
                    MessageBox.Show("Temporary geodatabase could not be created", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
            }

            // 2b. Create the empty output FC
            myArcMapFuncs.CreateFeatureClass(strOutFCName, strOutFolder, ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon, aLogFile);
            if (!myArcMapFuncs.FeatureclassExists(anOutputFile))
            {
                myFileFuncs.WriteLine(strLogFile, "Could not create output feature class " + anOutputFile);
                MessageBox.Show("Output feature class could not be created", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lngResult;
            }
            myFileFuncs.WriteLine(strLogFile, "Output Feature Class " + anOutputFile + " created.");

            // 2c. Add fields
            foreach (OutputColumn aCol in OutputLayer.OutputColumns)
            {
                // The used version of AddField takes account of simple column types.
                blTest = myArcMapFuncs.AddField(anOutputFile, aCol.ColumnName, aCol.FieldType, aCol.ColumnLength, aLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not create output field " + aCol.ColumnName);
                    MessageBox.Show("Field in output feature class could not be created", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
            }
            myFileFuncs.WriteLine(strLogFile, "New fields written to output Feature Class " + strOutFCName);

            // 3. For each input layer type (note we have points and polygons separate)
            foreach (MapLayer anInputLayer in InputLayers)
            {
                string strTempRawPoints = anInputLayer.LayerName;
                
                // If DissolveDistance > 0:
                if (anInputLayer.DissolveSize > 0)
                {
                    // FIRST OF ALL Work out what the clusters are.
                    // Copy the raw data to the temp directory
                    // 1a. Clear any selected features; make new selection if necessary.
                    myArcMapFuncs.ClearSelectedMapFeatures(anInputLayer.LayerName);
                    if (anInputLayer.WhereClause != "")
                    {
                        // Select.

                    }
                    

                    // 1b. Copy.
                    strTempRawPoints = strTempGDB + @"\TempRaw";
                    blTest = myArcMapFuncs.CopyFeatures(anInputLayer.LayerName, strTempRawPoints);
                    if (!blTest)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Could not create temporary feature class " + strTempRawPoints);
                        MessageBox.Show("Could not create temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }

                    // 2. Assign unique IDs
                    string strUniqueField = "HL_ID";
                    blTest =  myArcMapFuncs.AddField(strTempRawPoints, strUniqueField, "LONG", 10, strLogFile);
                    if (!blTest)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Could not add field " + strUniqueField + " to feature class " + strTempRawPoints);
                        MessageBox.Show("Could not add field to temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }
                    long intTest = myArcMapFuncs.AddIncrementalNumbers(strTempRawPoints, strUniqueField, aLogFile: strLogFile);
                    if (intTest <= 0)
                    {
                        myFileFuncs.WriteLine(strLogFile, "An error occurred while writing unique IDs to feature class " + strTempRawPoints);
                        MessageBox.Show("Error writing unique values to feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }

                    // 2. Buffer the points/polys with the DissolveDistance. Ignore aggregate fields here as we want all overlaps.
                    string strDissolveDistance = anInputLayer.DissolveSize.ToString() + " Meters";
                    string strBuffOutput = strTempGDB + "RawBuffered";
                    blTest = myArcMapFuncs.BufferFeatures(myFileFuncs.GetFileName(strTempRawPoints), strBuffOutput, strDissolveDistance, "", "NONE", strLogFile);
                    if (!blTest)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Error buffering temporary feature class " + strTempRawPoints + ".");
                        MessageBox.Show("Error buffering temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }

                    // 3. Dissolve the resulting polygons on key fields and overlap
                    string strDissolveOutput = strTempGDB + "BuffDissolved";

                    blTest = myArcMapFuncs.DissolveFeatures(myFileFuncs.GetFileName(strBuffOutput), strDissolveOutput, "Fieldlist", "", aLogFile: strLogFile);

                    // 4. Assign cluster IDs.
                    // 5. Spatial join of the original points/polys back onto this layer
                    // 6. Query where key fields are the same - this will give us the cluster IDs for each point/poly.
                    // 7. Calculate cluster ID back onto points/polys in new field using attribute join on unique ID.

                    // 8. Remove temporary layers
                    myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strTempRawPoints));
                    myArcMapFuncs.DeleteFeatureclass(strTempRawPoints); // Removes FC but the lock file remains.

                    myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strBuffOutput));
                    myArcMapFuncs.DeleteFeatureclass(strBuffOutput);

                    myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strDissolveOutput));
                    myArcMapFuncs.DeleteFeatureclass(strDissolveOutput);
                }
                else
                {
                    // 
                }
                
                // 8. We now have all the information that we need. Buffer all points/polys to required distance.
                // 9. Dissolve on key fields, including cluster ID if relevant. If DissolveDistance = 0 then Cluster fields are also key.
                // 10. Derive statistics during the dissolve, and for the common / cluster fields by multiple summaries to get max_count

                // 11. Dissolve taking into account key fields; if DissolveDistance = 0 then any Cluster fields are also a key field.
                // 12. Calculate any ranges as required using min and math range; keep track of field names or indices.

                // 13. Using an insert cursor add the new records to the empty FC.

                // 14. Remove all temporary layers.

                // 15. Delete all temporary files.

                // 16. Remember how many records we have added.

            }
            lngResult = 999;
            return lngResult; // Return the number of records inserted.
        }
    }
}
