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
using System.Collections;
using System.Linq;
using System.Text;

namespace DataBuffer
{
    // Full implementation of IEnumerable in this class in comments. Simplified version used.
    public class MapLayer
    {
        public string DisplayName { get; set; } // The name on the tool menu
        public string LayerName { get; set; } // the name in the TOC
        public InputColumns InputColumns { get; set; }
        public string WhereClause { get; set; }
        //public string SortOrder { get; set; }
        public int BufferSize { get; set; }
        public int DissolveSize { get; set; }

        // Constructor.
        public MapLayer()
        {   
        }
    }



    public class MapLayers : IEnumerable
    {
        private List<MapLayer> _layers;
        
        public MapLayers(MapLayer[] pArray = null) //
        {
            _layers = new List<MapLayer>(); //new MapLayer[pArray.Length];
            if (pArray != null)
            {
                for (int i = 0; i < pArray.Length; i++)
                {
                    _layers.Add(pArray[i]);
                }
            }
        }

        // Implementation of IEnumerable<>
        public IEnumerator<MapLayer> GetEnumerator()
        {
            foreach (MapLayer aLayer in _layers)
            {
                if (aLayer == null)
                {
                    break;
                }
                // Return the item
                yield return aLayer;
            }
        }

        // Implementation of IEnumerable
        IEnumerator IEnumerable.GetEnumerator() //public
        {
            return this.GetEnumerator();
        }


        //public LayerEnum GetEnumerator()
        //{
        //    return new LayerEnum(_layers);
        //}

        public void Add(MapLayer aLayer)
        {
            if (aLayer != null)
            {
                _layers.Add(aLayer);
            }
        }

        public MapLayer Get(int Index)
        {
            if (Index <= _layers.Count - 1)
            {
                return _layers[Index];
            }
            else
                return null;

        }
    }

    public class OutputLayer
    {
        public OutputColumns OutputColumns { get; set; }
        public string LayerPath { get; set; }
        public string LayerFile { get; set; }
        public string Format { get; set; }

        public OutputLayer()
        {
        }
    }

    public class InputColumn
    {
        public string InputName { get; set; }
        public string OutputName { get; set; }
        public string ColumnType { get; set; } // cluster, common, range etc.
        public string FieldType { get; set; } // int, double, text etc.
        public int FieldLength { get; set; }
    }

    public class InputColumns : IEnumerable
    {
        private List<InputColumn> _inputcolumns;
        public InputColumns(InputColumn[] pArray = null) // Constructor. We could use a list as input, too.
        {
            _inputcolumns = new List<InputColumn>();
            if (pArray != null)
            {
                for (int i = 0; i < pArray.Length; i++)
                {
                    _inputcolumns.Add(pArray[i]);
                }

            }
        }

        // Implement GetEnumerator
        public IEnumerator<InputColumn> GetEnumerator()
        {
            foreach (InputColumn aColumn in _inputcolumns)
            {
                if (aColumn == null)
                {
                    break;
                }
                // Return the item
                yield return aColumn;
            }
        }

        // Add method.
        public void Add(InputColumn anInputColumn)
        {
            if (anInputColumn != null)
            {
                _inputcolumns.Add(anInputColumn);
            }
        }

        // Implementation of generic GetEnumerator
        IEnumerator IEnumerable.GetEnumerator() //public
        {
            return this.GetEnumerator();
        }
    }

    public class OutputColumn
    {
        public string ColumnTag { get; set; } // Only used for error reporting.
        public string ColumnName { get; set; }
        public string ColumnType { get; set; }
        public string FieldType { get; set; }
        public int ColumnLength { get; set; }

        public OutputColumn()
        {
        }
    }

    public class OutputColumns : IEnumerable
    {
        private List<OutputColumn> _outputcolumns;
        public OutputColumns(OutputColumn[] pArray = null) // Constructor
        {
            _outputcolumns = new List<OutputColumn>();
            if (pArray != null)
            {
                for (int i = 0; i < pArray.Length; i++)
                {
                    _outputcolumns.Add(pArray[i]);
                }
            }
        }

        // Implement GetEnumerator
        public IEnumerator<OutputColumn> GetEnumerator()
        {
            foreach (OutputColumn aColumn in _outputcolumns)
            {
                if (aColumn == null)
                {
                    break;
                }
                // Return the item
                yield return aColumn;
            }
        }

        // Add method.
        public void Add(OutputColumn anOutputColumn)
        {
            if (anOutputColumn != null)
            {
                _outputcolumns.Add(anOutputColumn);
            }
        }

        // Implementation of generic GetEnumerator
        IEnumerator IEnumerable.GetEnumerator() //public
        {
            return this.GetEnumerator();
        }
    }

#region IEnumerator
    //// Must also implement IEnumerator IS THIS REALLY NECESSARY?? NO
    //public class LayerEnum : IEnumerator
    //{
    //    public MapLayer[] _layers;

    //    //Enumerators are positioned before the first element 
    //    // until the first MoveNext() call.
    //    int position = -1;

    //    public LayerEnum(MapLayer[] list)
    //    {
    //        _layers = list;
    //    }

    //    public bool MoveNext()
    //    {
    //        position++;
    //        return (position < _layers.Length);
    //    }

    //    public void Reset()
    //    {
    //        position = -1;
    //    }

    //    object IEnumerator.Current
    //    {
    //        get
    //        {
    //            return Current;
    //        }
    //    }

    //    public MapLayer Current
    //    {
    //        get
    //        {
    //            try
    //            {
    //                return _layers[position];
    //            }
    //            catch (IndexOutOfRangeException)
    //            {
    //                throw new InvalidOperationException();
    //            }
    //        }
    //    }
    //}
#endregion
}
