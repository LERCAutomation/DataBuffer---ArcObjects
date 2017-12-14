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

            long lngResult = 0; // The total number of rows written.
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
                myFileFuncs.WriteLine(strLogFile, "Temporary geodatabase created");
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
                string strTempRawLayer = "TempRaw";
                string strTempRawPoints = strTempGDB + @"\" + strTempRawLayer; ; // The layer we're going to do the calculations on 
                myFileFuncs.WriteLine(strLogFile, "Processing map layer " + anInputLayer.LayerName);

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
                blTest = myArcMapFuncs.CopyFeatures(anInputLayer.LayerName, strTempRawPoints);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not create temporary feature class " + strTempRawPoints);
                    MessageBox.Show("Could not create temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                myFileFuncs.WriteLine(strLogFile, "Temporary feature class created.");

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

                // FIRST OF ALL Work out what the clusters are. We do this for ALL CASES even if DissolveSize = 0.
                string strClusterIDField = "ClusterID";
                string strBuffOutput =  strTempRawPoints; // If we aren't dissolving on cluster, use the raw input.
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
                blTest = myArcMapFuncs.DissolveFeatures(myFileFuncs.GetFileName(strBuffOutput), strDissolveOutput, strFieldList, "", aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error dissolving temporary feature class " + strDissolveOutput + ".");
                    MessageBox.Show("Error dissolving temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 3. Assign cluster IDs.
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
                blTest = myArcMapFuncs.AddField(strTempRawPoints, strClusterIDField, "LONG", 10, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error adding Cluster ID field to temporary feature class " + strTempRawPoints + ".");
                    MessageBox.Show("Error adding new field to temporary feature class", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 5. Spatial join of the original points/polys back onto this layer. Note one-to-many join deals with overlaps.
                string strRawJoined = strTempGDB + @"\RawWithClusters";
                blTest = myArcMapFuncs.SpatialJoin(myFileFuncs.GetFileName(strTempRawPoints), myFileFuncs.GetFileName(strDissolveOutput), strRawJoined, aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error carrying out spatial join of " + strDissolveOutput + " to " + strTempRawPoints + ".");
                    MessageBox.Show("Error carrying out spatial join", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }
                string strRawJoinedLayer = myFileFuncs.GetFileName(strRawJoined); // the display name for the next function.

                // 6. Query where key fields are the same - this will give us the cluster IDs for each point/poly.
                // Build the query based on key columns.
                string strQuery = "";
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    if (aCol.ColumnType == "key" || (aCol.ColumnType=="cluster" && anInputLayer.DissolveSize == 0)) 
                    {
                        strQuery = strQuery + aCol.InputName + " = " + aCol.InputName + "_1 AND ";
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
                string strLookupTable = strTempGDB + @"\ClusterLookup";
                blTest = myArcMapFuncs.CopyFeatures(strRawJoinedLayer, strLookupTable, aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error exporting features from " + strRawJoinedLayer + ".");
                    MessageBox.Show("Error exporting features from temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 7. Calculate cluster ID back onto points/polys in new field using attribute join on unique ID.
                // 7a. Join to raw points. (temporary join)
                blTest = myArcMapFuncs.AddJoin(strTempRawLayer, strUniqueField, myFileFuncs.GetFileName(strLookupTable), strUniqueField, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error joining " + myFileFuncs.GetFileName(strLookupTable) + " to " + myFileFuncs.GetFileName(strTempRawPoints) + ".");
                    MessageBox.Show("Error joining lookup table to temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // 7b. Calculate
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
                myFileFuncs.WriteLine(strLogFile, "Cluster information transferred to temporary layer");

                // Now create the final temporary output layer for input into the Buffer layer.
                // Buffer all points/polys to required distance and dissolve on ClusterID. Even when DissolveSize =  0 we have a cluster ID.
                string strBufferredInput = strTempGDB + @"\FinalRawBuffered";
                string strBufferDistance = anInputLayer.BufferSize.ToString() + " Meters";
                string strDissolveField = strClusterIDField;
                string strDissolveOption = "LIST";
                

                blTest = myArcMapFuncs.BufferFeatures(myFileFuncs.GetFileName(strTempRawLayer), strBufferredInput, strBufferDistance, strDissolveField, strDissolveOption, aLogFile: strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error buffering input layer " + strTempRawLayer);
                    MessageBox.Show("Error buffering input layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                string strFinalInputLayer = myFileFuncs.GetFileName(strBufferredInput); // Note this has no relevant fields other than ClusterID.

                // Add all the fields.
                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
                    blTest = myArcMapFuncs.AddField(strBufferredInput, aCol.OutputName, aCol.FieldType, aCol.FieldLength, strLogFile);
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
                
                blTest = myArcMapFuncs.AddJoin(strFinalInputLayer, strClusterIDField, strStraightStats, strClusterIDField, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error joining temporary table " + myFileFuncs.GetFileName(strStraightStats) + " to temporary layer " + strFinalInputLayer);
                    MessageBox.Show("Error joining statistics table to temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                foreach (InputColumn aCol in anInputLayer.InputColumns)
                {
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
                }

                myArcMapFuncs.RemoveJoin(strFinalInputLayer, strLogFile);

                myFileFuncs.WriteLine(strLogFile, "first, min/max and range fields calculated.");

                // Now deal with the common and cluster fields.
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
                blTest = myArcMapFuncs.DeleteField(strFinalInputLayer, strClusterIDField, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Error deleting field from temporary layer " + strFinalInputLayer);
                    MessageBox.Show("Error deleting field from temporary layer", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Merge to create the final output layer.
                blTest = myArcMapFuncs.AppendFeatures(strFinalInputLayer, anOutputFile, strLogFile);
                if (!blTest)
                {
                    myFileFuncs.WriteLine(strLogFile, "Could not append temporary results " + strFinalInputLayer + " to output layer + " + anOutputFile);
                    MessageBox.Show("Error appending temporary results to output", "Data Buffer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return lngResult;
                }

                // Remember how many records we have added.
                long lngCount = myArcMapFuncs.CountAllLayerFeatures(strFinalInputLayer); // tidy up.
                if (lngCount > 0)
                    lngResult = lngResult + lngCount;
                myFileFuncs.WriteLine(strLogFile, "Results added to output layer. A total of " + lngCount.ToString() + " rows were added.");

                // Remove all temporary layers.
                // Delete all temporary files.

                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strTempRawPoints));
                myArcMapFuncs.DeleteFeatureclass(strTempRawPoints); // Removes FC but the lock file remains.

                if (myArcMapFuncs.FeatureclassExists(strDissolveOutput))
                {
                    myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strDissolveOutput));
                    myArcMapFuncs.DeleteFeatureclass(strDissolveOutput);
                }

                if (myArcMapFuncs.FeatureclassExists(strBuffOutput))
                {
                    myArcMapFuncs.RemoveStandaloneTable(myFileFuncs.GetFileName(strBuffOutput));
                    myArcMapFuncs.DeleteFeatureclass(strBuffOutput);
                }

                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strRawJoined));
                myArcMapFuncs.DeleteFeatureclass(strRawJoined);

                myArcMapFuncs.RemoveLayer(myFileFuncs.GetFileName(strLookupTable));
                myArcMapFuncs.DeleteFeatureclass(strLookupTable);

                myArcMapFuncs.RemoveLayer(strFinalInputLayer);
                myArcMapFuncs.DeleteFeatureclass(strBufferredInput);

                myArcMapFuncs.RemoveStandaloneTable(myFileFuncs.GetFileName(strStraightStats));
                myArcMapFuncs.DeleteFeatureclass(strStraightStats); // also deletes tables.

                myArcMapFuncs.ClearSelectedMapFeatures(anInputLayer.LayerName);

                myFileFuncs.WriteLine(strLogFile, "Temporary layers removed");
            }

            return lngResult; // Return the number of records inserted.
            myFileFuncs.WriteLine(strLogFile, "All layers processed.");
        }
    }
}
