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
            myFileFuncs.WriteLine(strLogFile, "The output layer is " + OutputFile + ".");
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

            // 3. For each species type (note we have points and polygons separate)

            // buffer to the required distance

            // Merge the two buffered layers as relevant.

            // We now have the rawest of inputs. 
            // 4. Dissolve on the selected key columns using as much as the innate functionality as possible
            //    KEEP A TRACK OF FIELD NAMES DURING ALL SUBSEQUENT PROCESSING
            //    Use Min and Range to allow for range calculations.
            //    This becomes the raw output file from which results will be read. 

            // 5. Calculate any ranges into new fields as required.

            // 6. To estimate modal value go through two sets of summary tables to get the max occurrence for 
            //    each common field for each key combo and use these during (7)

            // 7. Using an insert cursor add the new records to the empty FC

            return -999; // Return the number of records inserted.
        }
    }
}
