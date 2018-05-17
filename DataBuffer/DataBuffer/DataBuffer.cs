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
using System.Text;
using System.IO;

namespace DataBuffer
{
    public class DataBuffer : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public DataBuffer()
        {
        }

        protected override void OnClick()
        {
            frmDataBuffer myForm = new frmDataBuffer();
            myForm.ShowDialog();
            ArcMap.Application.CurrentTool = null;
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
