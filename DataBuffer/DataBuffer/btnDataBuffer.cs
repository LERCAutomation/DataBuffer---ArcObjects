using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DataBuffer
{
    public class btnDataBuffer : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public btnDataBuffer()
        {
        }

        protected override void OnClick()
        {
            frmDataBuffer myForm = new frmDataBuffer();
            myForm.ShowDialog();
            //
            //  TODO: Sample code showing how to access button host
            //
            ArcMap.Application.CurrentTool = null;
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
