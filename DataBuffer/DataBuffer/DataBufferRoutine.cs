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

        public long RunDataBuffer(string anOutputFile, string aLogFile, frmDataBuffer callingForm)
        {
            // This uses the settings XML to retrieve the layerfolder MAYBE
            // so they do not need to be passed as a setting.
            // Inputlayers, outputlayers and output file should already have been set.

            // Note this has been set up to be rather test-driven. While it is very hard to formally
            // unit test ArcGIS add-ins, this code attempts to be as airtight in terms of catching
            // all eventualities as I could get it.

            long lngResult = 0; // The total number of rows written.
            bool blTest = false; // Testing variable.

            callingForm.UpdateStatus("Starting process");

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

            callingForm.UpdateStatus("Creating output layer...", "Starting process");

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
                myFileFuncs.WriteLine(strLogFile, "Temporary geodatabase created");
            }


            // 2b. Create the empty output FC ini the temporary directory.


            callingForm.UpdateStatus(".");
            string strTempFinalOutputLayer = "TempFinal";
            string strTempFinalOutput = strTempGDB + @"\" + strTempFinalOutputLayer;
            myArcMapFuncs.CreateFeatureClassNew(strTempGDB, strTempFinalOutputLayer,"POLYGON", "", InputLayers.Get(0).LayerName, aLogFile);
            if (!myArcMapFuncs.FeatureclassExists(strTempFinalOutput))
            {
                myFileFuncs.WriteLine(strLogFile, "Could not create output feature class " + strTempFinalOutput);
                MessageBox.Show("Output feature class could not be created", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lngResult;
            }
            myFileFuncs.WriteLine(strLogFile, "Output Feature Class " + strTempFinalOutput + " created.");

            // 2c. Add fields
            foreach (OutputColumn aCol in OutputLayer.OutputColumns)
            {
                callingForm.UpdateStatus(".");
                // The used version of AddField takes account of simple column types.
                blTest = myArcMapFuncs.AddField(strTempFinalOutput, aCol.ColumnName, aCol.FieldType, aCol.ColumnLength, aLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not create output field " + aCol.ColumnName);
                    MessageBox.Show("Field in output feature class could not be created", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
            }
            myFileFuncs.WriteLine(strLogFile, "New fields written to output Feature Class " + strTempFinalOutput);
            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");

            // 3. For each input layer type (note we have points and polygons separate)
            foreach (MapLayer anInputLayer in InputLayers)
            {
                callingForm.UpdateStatus("Checking input", "Processing " + anInputLayer.DisplayName);
                
                // Firstly let's check the promised input fields exist and is of the correct type.
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    callingForm.UpdateStatus(".");
                    if (aCol.InputName.Substring(0, 1) != "\"")
                    {
                        if (!myArcMapFuncs.FieldExists(anInputLayer.LayerName, aCol.InputName, aLogFile: strLogFile))
                        {
                            myFileFuncs.WriteLine(strLogFile, "Cannot locate field " + aCol.InputName + " in layer " + anInputLayer.LayerName);
                            MessageBox.Show("Field " + aCol.InputName + " does not exist in layer " + anInputLayer.LayerName);
                            return lngResult;
                        }
                        else
                        {
                            if (!myArcMapFuncs.CheckFieldType(anInputLayer.LayerName, aCol.InputName, aCol.FieldType, strLogFile) && aCol.ColumnType != "range")
                            {
                                myFileFuncs.WriteLine(strLogFile, "The field " + aCol.InputName + " in layer " + anInputLayer.LayerName + " is not of type " + aCol.FieldType);
                                MessageBox.Show("Wrong field type for field " + aCol.InputName + " in layer " + anInputLayer.LayerName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    
                }

                string strTempRawLayer = "TempRaw";
                string strTempRawPoints = strTempGDB + @"\" + strTempRawLayer; ; // The layer we're going to do the calculations on 
                myFileFuncs.WriteLine(strLogFile, "Processing map layer " + anInputLayer.LayerName);

                callingForm.UpdateStatus("Copying to temporary file");
                // 1a. Clear any selected features; make new selection if necessary.
                myArcMapFuncs.ClearSelectedMapFeatures(anInputLayer.LayerName);
                if (anInputLayer.WhereClause != "")
                {
                    // Select.
                    blTest = myArcMapFuncs.SelectLayerByAttributes(anInputLayer.LayerName, anInputLayer.WhereClause, aLogFile: strLogFile);
                    if (!blTest)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Could not create selection on input layer " + anInputLayer.LayerName + " using where clause " + anInputLayer.WhereClause);
                        MessageBox.Show("Selection on input layer " + anInputLayer.LayerName + " failed", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }
                    int intSelCount = myArcMapFuncs.CountSelectedLayerFeatures(anInputLayer.LayerName);
                    if (intSelCount == 0)
                    {
                        // No features selected.
                        myFileFuncs.WriteLine(strLogFile, "No features selected.");
                        break;
                    }
                    else
                    {
                        myFileFuncs.WriteLine(strLogFile, intSelCount.ToString() + " features selected");
                    }
                }

                // Copy the raw data (inc selection) to the temp directory
                // 1b. Copy.
                callingForm.UpdateStatus(".");
                blTest = myArcMapFuncs.CopyFeatures(anInputLayer.LayerName, strTempRawPoints);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not create temporary feature class " + strTempRawPoints);
                    MessageBox.Show("Could not create temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                myFileFuncs.WriteLine(strLogFile, "Temporary feature class created.");

                // 2. Assign unique IDs
                callingForm.UpdateStatus(".");
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

                // Index the unique IDs
                blTest = myArcMapFuncs.CreateIndex(strTempRawPoints, strUniqueField, "TempRawIndex", strLogFile);
                if (!blTest) // Doesn't stop us carrying on.
                {
                    myFileFuncs.WriteLine(strLogFile, "WARNING: Could not add index to feature class " + strTempRawPoints + ". Process will be slow");
                }

                // FIRST OF ALL Work out what the clusters are. We do this for ALL CASES even if DissolveSize = 0.
                string strClusterIDField = "ClusterID";
                string strBuffOutput =  strTempRawPoints; // If we aren't dissolving on cluster, use the raw input.
                callingForm.UpdateStatus(".");
                if (anInputLayer.DissolveSize > 0)
                {
                    // 1. Buffer the points/polys with the DissolveDistance. Ignore aggregate fields here as we want all overlaps.
                    int intDissolveDistance = anInputLayer.DissolveSize;
                    if (intDissolveDistance == anInputLayer.BufferSize) intDissolveDistance--; // decrease by one to avoid merging adjacent buffers.
                    string strDissolveDistance = intDissolveDistance.ToString() + " Meters";
                    strBuffOutput = strTempGDB + @"\RawBuffered";
                    blTest = myArcMapFuncs.BufferFeatures(myFileFuncs.GetFileName(strTempRawPoints), strBuffOutput, strDissolveDistance, aLogFile: strLogFile);
                    if (!blTest)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Error buffering temporary feature class " + strTempRawPoints + ".");
                        MessageBox.Show("Error buffering temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }
                    myFileFuncs.WriteLine(strLogFile, "Clusters will be identified within a distance of " + intDissolveDistance.ToString() + " metres");
                }
                else
                {
                    myFileFuncs.WriteLine(strLogFile, "Clusters will be identified for records at the same location only");
                }
                // 2. Dissolve the resulting polygons on key fields and overlap
                callingForm.UpdateStatus("Dissolving");
                string strDissolveOutput = strTempGDB + @"\BuffDissolved";
                string strFieldList = "";
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                { // Note cluster fields become key fields if DissolveSize = 0
                    if (aCol.ColumnType == "key" || (aCol.ColumnType == "cluster" && anInputLayer.DissolveSize == 0)) 
                    {
                        strFieldList = strFieldList + aCol.InputName + ";";
                    }
                }
                strFieldList = strFieldList.Substring(0, strFieldList.Length - 1); // remove last ;

                // Index the fields first
                blTest = myArcMapFuncs.CreateIndex(strBuffOutput, strFieldList, "BuffOutputIndex", strLogFile);
                if (!blTest) // Doesn't stop us carrying on.
                {
                    myFileFuncs.WriteLine(strLogFile, "WARNING: Could not add key field index to feature class " + strBuffOutput + ". Process will be slow");
                }

                // Do the dissolve
                callingForm.UpdateStatus(".");
                blTest = myArcMapFuncs.DissolveFeatures(myFileFuncs.GetFileName(strBuffOutput), strDissolveOutput, strFieldList, "", aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error dissolving temporary feature class " + strDissolveOutput + ".");
                    MessageBox.Show("Error dissolving temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 3. Assign cluster IDs.
                callingForm.UpdateStatus(".");
                string strTempClusterIDField = "tClusterID";
                blTest = myArcMapFuncs.AddField(strDissolveOutput,strTempClusterIDField, "LONG", 10, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error adding Cluster ID field to temporary feature class " + strDissolveOutput + ".");
                    MessageBox.Show("Error adding new field to temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                intTest = myArcMapFuncs.AddIncrementalNumbers(strDissolveOutput, strTempClusterIDField, 1, strLogFile);
                if (intTest <= 0)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error adding Cluster ID to temporary feature class " + strDissolveOutput + ".");
                    MessageBox.Show("Error adding cluster IDs to temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                myFileFuncs.WriteLine(strLogFile, "Cluster numbers assigned.");

                // 4. Add Cluster ID field to raw points.
                callingForm.UpdateStatus(".");
                blTest = myArcMapFuncs.AddField(strTempRawPoints, strClusterIDField, "LONG", 10, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error adding Cluster ID field to temporary feature class " + strTempRawPoints + ".");
                    MessageBox.Show("Error adding new field to temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 5. Spatial join of the original points/polys back onto this layer. Note one-to-many join deals with overlaps.
                callingForm.UpdateStatus("Calculating clusters");
                
                string strRawJoined = strTempGDB + @"\RawWithClusters";
                blTest = myArcMapFuncs.SpatialJoin(myFileFuncs.GetFileName(strTempRawPoints), myFileFuncs.GetFileName(strDissolveOutput), strRawJoined, aMatchMethod: "INTERSECT", aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error carrying out spatial join of " + strDissolveOutput + " to " + strTempRawPoints + ".");
                    MessageBox.Show("Error carrying out spatial join", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                string strRawJoinedLayer = myFileFuncs.GetFileName(strRawJoined); // the display name for the next function.

                // 6. Query where key fields are the same - this will give us the cluster IDs for each point/poly.
                // Build the query based on key columns.
                callingForm.UpdateStatus(".");
                List<string> theCriteriaFields = new List<string>();
                string strQuery = "";
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    if (aCol.ColumnType == "key" || (aCol.ColumnType=="cluster" && anInputLayer.DissolveSize == 0)) 
                    {
                        strQuery = strQuery + aCol.InputName + " = " + aCol.InputName + "_1 AND ";
                        theCriteriaFields.Add(aCol.InputName);
                    }
                }
                // Remove the last AND
                strQuery = strQuery.Substring(0, strQuery.Length - 5);
                blTest = myArcMapFuncs.SelectLayerByAttributes(strRawJoinedLayer, strQuery, aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error selecting features from " + strRawJoinedLayer + ".");
                    MessageBox.Show("Error selecting from temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Export this selection
                callingForm.UpdateStatus(".");
                string strLookupTable = strTempGDB + @"\ClusterLookup";
                blTest = myArcMapFuncs.CopyFeatures(strRawJoinedLayer, strLookupTable, aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error exporting features from " + strRawJoinedLayer + ".");
                    MessageBox.Show("Error exporting features from temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Index the unique ID in the lookup table.
                blTest = myArcMapFuncs.CreateIndex(strLookupTable, strUniqueField, "LookupIndex", strLogFile);
                if (!blTest) // Doesn't stop us carrying on.
                {
                    myFileFuncs.WriteLine(strLogFile, "WARNING: Could not add index to feature class " + strLookupTable + ". Process will be slow");
                }

                // 7. Calculate cluster ID back onto points/polys in new field using attribute join on unique ID.
                // 7a. Join to raw points. (temporary join)
                callingForm.UpdateStatus(".");
                blTest = myArcMapFuncs.AddJoin(strTempRawLayer, strUniqueField, myFileFuncs.GetFileName(strLookupTable), strUniqueField, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error joining " + myFileFuncs.GetFileName(strLookupTable) + " to " + myFileFuncs.GetFileName(strTempRawPoints) + ".");
                    MessageBox.Show("Error joining lookup table to temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 7b. Calculate
                callingForm.UpdateStatus(".");
                string strCalc =  "[" + myFileFuncs.GetFileName(strLookupTable) + "." + strTempClusterIDField + "]";
                blTest = myArcMapFuncs.CalculateField(strTempRawLayer, strTempRawLayer + "." + strClusterIDField, strCalc, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error calculating field " + strTempRawLayer + "." + strClusterIDField + " in  layer" + strTempRawLayer);                      
                    MessageBox.Show("Error calculating clusterID field in temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 7c. Remove join
                blTest = myArcMapFuncs.RemoveJoin(strTempRawLayer, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error removing join in " + strTempRawLayer);
                    MessageBox.Show("Error removing join in temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Sort out the cases where the join hasn't worked. This addresses a very odd bug in ArcGIS.

                intTest = myArcMapFuncs.SetValueFromUnderlyingLayer(strTempRawLayer, strUniqueField, new List<string>() { strClusterIDField }, theCriteriaFields, myFileFuncs.GetFileName(strDissolveOutput), new List<string>() { strTempClusterIDField }, aLogFile);
                if (intTest < 0)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not assign missing values to " + strTempRawLayer);
                    MessageBox.Show("Could not fix missing cluster IDs in temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                else if (intTest == 0)
                {
                    myFileFuncs.WriteLine(strLogFile, "There were no missing cluster IDs encountered for layer " + anInputLayer.LayerName);
                }
                else
                {
                    myFileFuncs.WriteLine(strLogFile, intTest.ToString() + " missing cluster IDs fixed for layer " + anInputLayer.LayerName);
                }
                
                // Index on ClusterID.
                blTest = myArcMapFuncs.CreateIndex(strTempRawPoints, strClusterIDField, "RawClusterIndex", strLogFile);
                if (!blTest) // Doesn't stop us carrying on.
                {
                    myFileFuncs.WriteLine(strLogFile, "WARNING: Could not add cluster index to feature class " + strTempRawPoints + ". Process will be slow");
                }

                myFileFuncs.WriteLine(strLogFile, "Cluster information transferred to temporary layer");

                // Now create the final temporary output layer for input into the Buffer layer.
                // Buffer all points/polys to required distance and dissolve on ClusterID. Even when DissolveSize =  0 we have a cluster ID.
                string strBufferedInput = strTempGDB + @"\FinalRawBuffered";
                string strDissolvedInput = strTempGDB + @"\FinalRawDissolved";
                string strBufferDistance = anInputLayer.BufferSize.ToString() + " Meters";
                string strDissolveField = strClusterIDField;
                //string strDissolveOption = "LIST";

                callingForm.UpdateStatus("Buffering Features");

                strQuery = strUniqueField + " > 0"; // for some reason we're not buffering anything - try and fix.
                blTest = myArcMapFuncs.SelectLayerByAttributes(strTempRawLayer, strQuery, aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not select from raw input " + strTempRawLayer);
                    MessageBox.Show("Unable to select features in temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                if (myArcMapFuncs.CountSelectedLayerFeatures(strTempRawLayer) == 0)
                {
                    myFileFuncs.WriteLine(strLogFile, "No features selected for buffering on temporary layer " + strTempRawLayer);
                    MessageBox.Show("No features selected on temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                // It is considerably quicker to run buffer then dissolve, rather than dissolve during buffer operation.
                blTest = myArcMapFuncs.BufferFeatures(strTempRawLayer, strBufferedInput, strBufferDistance, aLogFile: strLogFile); // No dissolve.
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error buffering input layer " + strTempRawLayer);
                    MessageBox.Show("Error buffering input layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Index the ClusterID
                callingForm.UpdateStatus(".");
                blTest = myArcMapFuncs.CreateIndex(strBufferedInput, strClusterIDField, "BuffInputIndex", strLogFile);
                if (!blTest) // Doesn't stop us carrying on.
                {
                    myFileFuncs.WriteLine(strLogFile, "WARNING: Could not add index to feature class " + strBufferedInput + ". Process will be slow");
                }
                callingForm.UpdateStatus(".");
                blTest = myArcMapFuncs.DissolveFeatures(myFileFuncs.GetFileName(strBufferedInput), strDissolvedInput, strDissolveField, aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error dissolving temporary layer " + strBufferedInput);
                    MessageBox.Show("Error dissolving temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                string strFinalInputLayer = myFileFuncs.GetFileName(strDissolvedInput); // Note this has no relevant fields other than ClusterID.

                // Add all the fields.
                callingForm.UpdateStatus(".");
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    blTest = myArcMapFuncs.AddField(strDissolvedInput, aCol.OutputName, aCol.FieldType, aCol.FieldLength, strLogFile);
                    if (!blTest)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Error adding field " + aCol.OutputName + " to " + strFinalInputLayer);
                        MessageBox.Show("Error adding field to temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }
                }

                // Get the summary statistics for everything but the common/cluster fields.
                string strStatsFields = "";
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    // filter out text entry columns.
                    if (aCol.InputName.Substring(0, 1) != "\"")
                    {
                        if (aCol.ColumnType == "key" || aCol.ColumnType == "first" || (aCol.ColumnType == "cluster" && anInputLayer.DissolveSize == 0))
                        {
                            strStatsFields = strStatsFields + aCol.InputName + " FIRST;";
                        }
                        else if (aCol.ColumnType == "min")
                        {
                            strStatsFields = strStatsFields + aCol.InputName + " MIN;";
                        }
                        else if (aCol.ColumnType == "max")
                        {
                            strStatsFields = strStatsFields + aCol.InputName + " MAX;";
                        }
                        else if (aCol.ColumnType == "range") // special case
                        {
                            strStatsFields = strStatsFields + aCol.InputName + " MIN;";
                            strStatsFields = strStatsFields + aCol.InputName + " MAX;";
                        }
                    }
                }
                strStatsFields = strStatsFields.Substring(0, strStatsFields.Length - 1); // remove last semicolon.
                callingForm.UpdateStatus(".");
                string strStraightStatsLayer = "Straightstats";
                string strStraightStats = strTempGDB + @"\"+ strStraightStatsLayer;
                blTest = myArcMapFuncs.SummaryStatistics(strTempRawLayer, strStraightStats, strStatsFields, strClusterIDField, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error summarising temporary layer " + strTempRawLayer + " for initial statistics");
                    MessageBox.Show("Error summarising temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Join back to the final input layer, and calculate into new fields.
                callingForm.UpdateStatus("Calculating results");
                
                blTest = myArcMapFuncs.AddJoin(strFinalInputLayer, strClusterIDField, strStraightStats, strClusterIDField, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error joining temporary table " + myFileFuncs.GetFileName(strStraightStats) + " to temporary layer " + strFinalInputLayer);
                    MessageBox.Show("Error joining statistics table to temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    callingForm.UpdateStatus(".");
                    string strTargetField = strFinalInputLayer + "." + aCol.OutputName;
                    if (aCol.InputName.Substring(0, 1) == "\"")
                    {
                        strCalc = aCol.InputName;
                    }
                    else if (aCol.ColumnType == "key" || aCol.ColumnType == "first" || (aCol.ColumnType == "cluster" && anInputLayer.DissolveSize == 0))
                    {
                        strCalc = "[" + strStraightStatsLayer + "." + "FIRST_" + aCol.InputName + "]";
                    }
                    else if (aCol.ColumnType == "min")
                    {
                        strCalc = "[" + strStraightStatsLayer + "." + "MIN_" + aCol.InputName + "]";
                    }
                    else if (aCol.ColumnType == "max")
                    {
                        strCalc = "[" + strStraightStatsLayer + "." + "MAX_" + aCol.InputName + "]";
                    }
                    else if (aCol.ColumnType == "range")
                    {
                        // range: "min-max"
                        strCalc = "[" + strStraightStatsLayer + "." + "MIN_" + aCol.InputName + "] & \"-\" &";
                        strCalc = strCalc + "[" + strStraightStatsLayer + "." + "MAX_" + aCol.InputName + "]";
                    }
                    blTest = myArcMapFuncs.CalculateField(strFinalInputLayer, aCol.OutputName, strCalc, strLogFile);
                    if (!blTest)
                    {
                        myFileFuncs.WriteLine(strLogFile, "Error calculating field " + aCol.OutputName + " in temporary layer " + strFinalInputLayer + " using the following calculation: " + strCalc);
                        MessageBox.Show("Error calculating into temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return lngResult;
                    }
                    // Fix the special case where the column type is 'range' and the range covers a single value.
                    if (aCol.ColumnType == "range")
                    {
                        strQuery = strStraightStatsLayer + "." + "MIN_" + aCol.InputName + " = " + strStraightStatsLayer + "." + "MAX_" + aCol.InputName;
                        blTest = myArcMapFuncs.SelectLayerByAttributes(strFinalInputLayer, strQuery, aLogFile: strLogFile);
                        if (!blTest)
                        {
                            myFileFuncs.WriteLine(strLogFile, "Error selecting features from " + strFinalInputLayer + " using selection query " + strQuery);
                            MessageBox.Show("Error selecting features from " + strFinalInputLayer, "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return lngResult;
                        }
                        if (myArcMapFuncs.CountSelectedLayerFeatures(strFinalInputLayer) > 0)
                        {
                            strCalc = "[" + strStraightStatsLayer + "." + "MIN_" + aCol.InputName + "]";
                            blTest = myArcMapFuncs.CalculateField(strFinalInputLayer, aCol.OutputName, strCalc, strLogFile);
                            if (!blTest)
                            {
                                myFileFuncs.WriteLine(strLogFile, "Error calculating field " + aCol.OutputName + " in temporary layer " + strFinalInputLayer + " using the following calculation: " + strCalc);
                                MessageBox.Show("Error calculating into temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return lngResult;
                            }

                            myArcMapFuncs.ClearSelectedMapFeatures(strFinalInputLayer);
                        }
                    }

                }

                myArcMapFuncs.RemoveJoin(strFinalInputLayer, strLogFile);

                myFileFuncs.WriteLine(strLogFile, "first, min/max and range fields calculated.");

                // Now deal with the common and cluster fields.
                callingForm.UpdateStatus(".");
                List<string> strCommonInputFields = new List<string>();
                List<string> strCommonOutputFields = new List<string>();
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    if (aCol.ColumnType == "common" || (aCol.ColumnType == "cluster" && anInputLayer.DissolveSize > 0))
                    {
                        strCommonInputFields.Add(aCol.InputName);
                        strCommonOutputFields.Add(aCol.OutputName);
                    }
                }
                blTest = myArcMapFuncs.SetMostCommon(strFinalInputLayer, strClusterIDField, strCommonOutputFields, strTempRawLayer, strCommonInputFields, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error calculating common fields in temporary layer " + strFinalInputLayer);
                    MessageBox.Show("Error calculating common fields into temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                myFileFuncs.WriteLine(strLogFile, "Common and cluster values calculated.");

                // Drop the ClusterID field from the output
                callingForm.UpdateStatus(".");
                blTest = myArcMapFuncs.DeleteField(strFinalInputLayer, strClusterIDField, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error deleting field from temporary layer " + strFinalInputLayer);
                    MessageBox.Show("Error deleting field from temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Merge to create the final output layer.
                callingForm.UpdateStatus("Saving output");
                blTest = myArcMapFuncs.AppendFeatures(strFinalInputLayer, strTempFinalOutput, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not append temporary results " + strFinalInputLayer + " to output layer + " + anOutputFile);
                    MessageBox.Show("Error appending temporary results to output", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Remember how many records we have added.
                callingForm.UpdateStatus(".");
                long lngCount = myArcMapFuncs.CountAllLayerFeatures(strFinalInputLayer); // tidy up.
                if (lngCount > 0)
                    lngResult = lngResult + lngCount;
                myFileFuncs.WriteLine(strLogFile, "Results added to output layer. A total of " + lngCount.ToString() + " rows were added.");
                myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");

                // Remove all temporary layers.
                // Delete all temporary files.

                callingForm.UpdateStatus(".");
                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strTempRawPoints));
                myArcMapFuncs.DeleteFeatureclass(strTempRawPoints); // Removes FC but the lock file remains.

                if (myArcMapFuncs.FeatureclassExists(strDissolveOutput))
                {
                    myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strDissolveOutput));
                    myArcMapFuncs.DeleteFeatureclass(strDissolveOutput);
                }
                if (myArcMapFuncs.FeatureclassExists(strBuffOutput))
                {
                    myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strBuffOutput));
                    myArcMapFuncs.DeleteFeatureclass(strBuffOutput);
                }

                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strRawJoined));
                myArcMapFuncs.DeleteFeatureclass(strRawJoined);

                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strLookupTable));
                myArcMapFuncs.DeleteFeatureclass(strLookupTable);

                callingForm.UpdateStatus(".");
                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strDissolvedInput));
                myArcMapFuncs.DeleteFeatureclass(strDissolvedInput);

                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strBufferedInput));
                myArcMapFuncs.DeleteFeatureclass(strBufferedInput);

                myArcMapFuncs.RemoveLayer(strTempFinalOutputLayer);
                myArcMapFuncs.DeleteFeatureclass(strTempFinalOutput);

                myArcMapFuncs.RemoveStandaloneTable(myFileFuncs.GetFileName(strStraightStats));
                myArcMapFuncs.DeleteFeatureclass(strStraightStats); // also deletes tables.

                myArcMapFuncs.ClearSelectedMapFeatures(anInputLayer.LayerName);

                myFileFuncs.WriteLine(strLogFile, "Temporary layers removed");
            }
            myFileFuncs.WriteLine(strLogFile, "All layers processed.");

            // Write the final output.
            blTest = myArcMapFuncs.CopyFeatures(strTempFinalOutput, anOutputFile, aLogFile: strLogFile);
            if (!blTest)
            {
                myFileFuncs.WriteLine(strLogFile, "Could not copy temporary results " + strTempFinalOutput + " to final output layer + " + anOutputFile);
                MessageBox.Show("Error copy temporary results to final output", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return lngResult;
            }

            // Set the legend
            string strOutLayer = myFileFuncs.GetFileName(anOutputFile);
            if (strOutLayer.Substring(strOutLayer.Length - 4, 4) == ".shp")
                strOutLayer = myFileFuncs.ReturnWithoutExtension(strOutLayer);
            myArcMapFuncs.ChangeLegend(strOutLayer, OutputLayer.LayerFile, aLogFile: strLogFile);

            myArcMapFuncs.SetContentsView();
            callingForm.UpdateStatus("", "");
            
            return lngResult; // Return the number of records inserted.
        }
    }
}
