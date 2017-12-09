using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using HLDataBufferConfig;

namespace DataBuffer
{
    public partial class frmDataBuffer : Form
    {
        public frmDataBuffer()
        {
            // Firstly let's read the XML.
            DataBufferConfig myConfig = new DataBufferConfig();
            InitializeComponent();
            
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
