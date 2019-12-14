using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReporteComercial.Forms
{
    public partial class VisorDeReportes : Form
    {
        SqlConnection conexion;
        SqlCommand command;
        SqlDataAdapter dataAdapter;
        

        public VisorDeReportes()
        {
            InitializeComponent();
        }

        public VisorDeReportes(SqlConnection conexion)
        {
            InitializeComponent();
            this.conexion = conexion;
        }

        private void VisorDeReportes_Load(object sender, EventArgs e)
        {
            //this.reportViewer2.RefreshReport();
        }


        public void ReporteVentas(string fechaInicial, string fechaFinal)
        {
           




        }

    }
}
