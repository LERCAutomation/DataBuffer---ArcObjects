using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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

            // Check input. None of these should ever be fired.
            if (InputLayers == null)
            {
                MessageBox.Show("The input layers have not been defined.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -999;
            }
            else if ( OutputLayer == null)
            {
                MessageBox.Show("The output layer has not been defined.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -999;
            }
            else if (OutputFile == "")
            {
                MessageBox.Show("The output file has not been defined.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -999;
            }

            // All good to go.

            // 1. Start the log file. 
            string strLogFile = aLogFile;
            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");
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

            myArcMapFuncs.ToggleDrawing(false);
            myArcMapFuncs.ToggleTOC(false);

            // 2. Create the empty output FC using correct definitions.
            string strOutFCName = myFileFuncs.GetFileName(anOutputFile);
            string strOutFolder = myFileFuncs.GetDirectoryName(anOutputFile);
            // 2a. Create empty FC
            myArcMapFuncs.CreateFeatureClass(strOutFCName, strOutFolder, ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon, aLogFile);
            myFileFuncs.WriteLine(strLogFile, "Output Feature Class " + anOutputFile + " created.");
            // 2b. Add fields
            foreach (OutputColumn aCol in OutputLayer.OutputColumns)
            {
                // What is the column type?

                myArcMapFuncs.AddField(anOutputFile, aCol.ColumnName, aCol.FieldType, aCol.ColumnLength, aLogFile);
            }
            myFileFuncs.WriteLine(strLogFile, "New fields written to output Feature Class " + strOutFCName);

            // 3. For each input layer type (note we have points and polygons separate)
            foreach (MapLayer anInputLayer in InputLayers)
            {

                // If DissolveDistance > 0:
                // FIRST OF ALL Work out what the clusters are.
                // 1. Assign unique IDs
                // 2. Buffer the points/polys with the DissolveDistance.
                // 3. Dissolve the resulting polygons on key fields and overlap
                // 4. Assign cluster IDs.
                // 5. Spatial join of the original points/polys back onto this layer
                // 6. Query where key fields are the same - this will give us the cluster IDs for each point/poly.
                // 7. Calculate cluster ID back onto points/polys in new field using attribute join on unique ID.
                
                
                // 8. We now have all the information that we need. Buffer all points/polys to required distance.
                // 9. Dissolve on key fields, including cluster ID if relevant. If DissolveDistance = 0 then Cluster fields are also key.
                // 10. Derive statistics during the dissolve, and for the common / cluster fields by multiple summaries to get max_count

                // 11. Dissolve taking into account key fields; if DissolveDistance = 0 then any Cluster fields are also a key field.
                // 12. Calculate any ranges as required using min and math range; keep track of field names or indices.

                // 13. Using an insert cursor add the new records to the empty FC.
                

            }
            myArcMapFuncs.ToggleDrawing(true);
            myArcMapFuncs.ToggleTOC(true);
            return -999; // Return the number of records inserted.
        }
    }
}
