using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReporteComercial.Forms
{
    public partial class VisorDeReportes : Form
    {
        public VisorDeReportes()
        {
            InitializeComponent();
        }

        private void VisorDeReportes_Load(object sender, EventArgs e)
        {

            this.reportViewer2.RefreshReport();
        }
    }
}
