using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace cost_management
{
    public partial class Frmloading : Form
    {
        public Frmloading()
        {
            InitializeComponent();
        }

        private void Frmloading_Load(object sender, EventArgs e)
        {
            PBWait.Visible = true;
        }
    }
}
