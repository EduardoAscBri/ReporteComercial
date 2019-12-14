using ClosedXML.Excel;
using ReporteComercial.Classes;
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
    public partial class ReporteVentas : Form
    {
        string rutaLocal;
        SqlConnection conexion;

        public ReporteVentas()
        {
            InitializeComponent();
        }

        private void ReporteVentas_Load(object sender, EventArgs e)
        {
            this.rutaLocal = Application.StartupPath;
            InnerDatabaseManager databaseManager = new InnerDatabaseManager();
            this.conexion = databaseManager.createConnectionFromIniFile(this.rutaLocal);
        }

        private void bttSalir_Click(object sender, EventArgs e)
        {
            this.conexion.Close();
            this.Dispose();
        }

        private void bttAcept_Click(object sender, EventArgs e)
        {
            string lFechaInicial = "";
            string lFechaFinal = "";

            lFechaInicial = this.dtpFechaInicial.Text;
            lFechaFinal = this.dtpFechaFinal.Text;

            //lFechaInicial = "01/10/2019";
            //lFechaFinal = "30/10/2019";

            reporte(lFechaInicial, lFechaFinal);

        }

        public void reporte(string fechaInicial, string fechaFinal)
        {
            DataTable reporte = new DataTable();
            reporte.Columns.Add("CSERIEDOCUMENTO");
            reporte.Columns.Add("CFOLIO");
            reporte.Columns.Add("CFECHA");
            reporte.Columns.Add("CCODIGOCLIENTE");
            reporte.Columns.Add("CRAZONSOCIAL");
            reporte.Columns.Add("CRFC");
            reporte.Columns.Add("CTEXTOEXTRA1");
            reporte.Columns.Add("CTEXTOEXTRA2");
            reporte.Columns.Add("CTEXTOEXTRA3");
            reporte.Columns.Add("CPAIS");
            reporte.Columns.Add("CESTADO");
            reporte.Columns.Add("CCIUDAD");
            reporte.Columns.Add("CMUNICIPIO");
            reporte.Columns.Add("CCODIGOPOSTAL");
            reporte.Columns.Add("CCOLONIA");
            reporte.Columns.Add("CNUMEROINTERIOR");
            reporte.Columns.Add("CNUMEROEXTERIOR");
            reporte.Columns.Add("CNOMBRECALLE");
            reporte.Columns.Add("CTELEFONO1");
            reporte.Columns.Add("CEMAIL");
            reporte.Columns.Add("TIPOCLIENTE");
            reporte.Columns.Add("ZONA");
            reporte.Columns.Add("CCODIGOAGENTE");
            reporte.Columns.Add("CNOMBREAGENTE");
            reporte.Columns.Add("CCODIGOPRODUCTO");
            reporte.Columns.Add("CNOMBREPRODUCTO");
            reporte.Columns.Add("FAMILIA");
            reporte.Columns.Add("SISTEMA");
            reporte.Columns.Add("TIPO");
            reporte.Columns.Add("CNUMEROMOVIMIENTO");
            reporte.Columns.Add("CUNIDADES");
            reporte.Columns.Add("CPRECIO");
            reporte.Columns.Add("CNETO");
            reporte.Columns.Add("CDESCUENTO1");
            reporte.Columns.Add("CDESCUENTO2");
            reporte.Columns.Add("CTOTAL");
            reporte.Columns.Add("CREFERENCIA");
            reporte.Columns.Add("COBSERVAMOV");
            reporte.Columns.Add("SERIEPAGO");
            reporte.Columns.Add("FOLIOPAGO");
            reporte.Columns.Add("FECHAPAGO");
            reporte.Columns.Add("IMPORTEPAGADO");

            string cmdFacturas = "SELECT CIDDOCUMENTO, CSERIEDOCUMENTO, CFOLIO, CFECHA " +
                "FROM admDocumentos " +
                "WHERE(admDocumentos.CIDDOCUMENTODE = 4) " +
                "AND (admDocumentos.CFECHA BETWEEN @fechaInicial AND @fechaFinal)";

            string cmdFacturaDetalle = "SELECT        admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, " +
                "admClientes.CCODIGOCLIENTE, admClientes.CRAZONSOCIAL, admClientes.CRFC, " +
                "admClientes.CTEXTOEXTRA1, admClientes.CTEXTOEXTRA2, admClientes.CTEXTOEXTRA3, " +
                "admDomicilios.CPAIS, admDomicilios.CESTADO, admDomicilios.CCIUDAD, admDomicilios.CMUNICIPIO, admDomicilios.CCODIGOPOSTAL, admDomicilios.CCOLONIA, " +
                "admDomicilios.CNUMEROINTERIOR, admDomicilios.CNUMEROEXTERIOR, admDomicilios.CNOMBRECALLE, admDomicilios.CTELEFONO1, admDomicilios.CEMAIL, " +
                "admClasificacionesValores_1.CVALORCLASIFICACION AS TIPOCLIENTE, admClasificacionesValores_2.CVALORCLASIFICACION AS ZONA, " +
                "admAgentes.CCODIGOAGENTE, admAgentes.CNOMBREAGENTE, " +
                "admProductos.CCODIGOPRODUCTO, admProductos.CNOMBREPRODUCTO, " +
                "admClasificacionesValores_3.CVALORCLASIFICACION AS FAMILIA, " +
                "admClasificacionesValores_4.CVALORCLASIFICACION AS SISTEMA, " +
                "admClasificacionesValores_5.CVALORCLASIFICACION AS TIPO, " +
                "admMovimientos.CNUMEROMOVIMIENTO, admMovimientos.CUNIDADES, admMovimientos.CPRECIO, admMovimientos.CNETO, " +
                "admMovimientos.CDESCUENTO1, admMovimientos.CDESCUENTO2, " +
                "admMovimientos.CTOTAL, " +
                "admMovimientos.CREFERENCIA, admMovimientos.COBSERVAMOV " +
                "FROM admDocumentos " +
                "LEFT JOIN admClientes " +
                "ON admDocumentos.CIDCLIENTEPROVEEDOR = admClientes.CIDCLIENTEPROVEEDOR " +
                "LEFT JOIN admDomicilios " +
                "ON admClientes.CIDCLIENTEPROVEEDOR = admDomicilios.CIDCATALOGO " +
                "LEFT JOIN admClasificacionesValores AS admClasificacionesValores_1 " +
                "ON admClientes.CIDVALORCLASIFCLIENTE1 = admClasificacionesValores_1.CIDVALORCLASIFICACION " +
                "LEFT JOIN admClasificacionesValores AS admClasificacionesValores_2 " +
                "ON admClientes.CIDVALORCLASIFCLIENTE2 = admClasificacionesValores_2.CIDVALORCLASIFICACION " +
                "LEFT JOIN admAgentes " +
                "ON admDocumentos.CIDAGENTE = admAgentes.CIDAGENTE " +
                "LEFT JOIN admMovimientos " +
                "ON admDocumentos.CIDDOCUMENTO = admMovimientos.CIDDOCUMENTO " +
                "LEFT JOIN admProductos " +
                "ON admMovimientos.CIDPRODUCTO = admProductos.CIDPRODUCTO " +
                "LEFT JOIN admClasificacionesValores AS admClasificacionesValores_3 " +
                "ON admProductos.CIDVALORCLASIFICACION1 = admClasificacionesValores_3.CIDVALORCLASIFICACION " +
                "LEFT JOIN admClasificacionesValores AS admClasificacionesValores_4 " +
                "ON admProductos.CIDVALORCLASIFICACION4 = admClasificacionesValores_4.CIDVALORCLASIFICACION " +
                "LEFT JOIN admClasificacionesValores AS admClasificacionesValores_5 " +
                "ON admProductos.CIDVALORCLASIFICACION5 = admClasificacionesValores_5.CIDVALORCLASIFICACION " +
                "WHERE(admDocumentos.CIDDOCUMENTODE = 4) AND (admDocumentos.CIDDOCUMENTO = @cidfactura) AND (admDomicilios.CTIPOCATALOGO = 1) AND(admDomicilios.CTIPODIRECCION = 0) " +
                "ORDER BY admDocumentos.CFOLIO ASC";

            string cmdCargosAbonos = "SELECT        admAsocCargosAbonos.CIDDOCUMENTOABONO, admAsocCargosAbonos.CIDDOCUMENTOCARGO, " +
                "admAsocCargosAbonos.CIMPORTEABONO, admAsocCargosAbonos.CIMPORTECARGO, " +
                "admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, admDocumentos.CTOTAL " +
                "FROM admAsocCargosAbonos " +
                "LEFT JOIN admDocumentos " +
                "ON admAsocCargosAbonos.CIDDOCUMENTOABONO = admDocumentos.CIDDOCUMENTO " +
                "WHERE(admDocumentos.CIDDOCUMENTODE = 9) AND (admAsocCargosAbonos.CIDDOCUMENTOCARGO = @cidfactura) AND (admDocumentos.CFECHA BETWEEN @fechaInicial AND @fechaFinal)";

            SqlCommand command = new SqlCommand();
            command = new SqlCommand(cmdFacturas, this.conexion);
            
            command.Parameters.AddWithValue("@fechaInicial", fechaInicial);
            command.Parameters.AddWithValue("@fechaFinal", fechaFinal);
            
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
            DataTable listaFacturas = new DataTable();
            dataAdapter.Fill(listaFacturas);

            DataTable detalleFactura = new DataTable();
            DataTable detallePago = new DataTable();

            

            foreach(DataRow row in listaFacturas.Rows)
            {
                int cidfactura = Convert.ToInt32(row["CIDDOCUMENTO"]);

                //Consulta detalle de la factura
                detalleFactura.Clear();
                command = new SqlCommand(cmdFacturaDetalle, this.conexion);
                command.Parameters.AddWithValue("@cidfactura", cidfactura);
                dataAdapter = new SqlDataAdapter(command);
                dataAdapter.Fill(detalleFactura);

                //Consulta pagos aplicados
                detallePago.Clear();
                command = new SqlCommand(cmdCargosAbonos, this.conexion);
                command.Parameters.AddWithValue("@cidfactura", cidfactura);
                command.Parameters.AddWithValue("@fechaInicial", fechaInicial);
                command.Parameters.AddWithValue("@fechaFinal", fechaFinal);
                dataAdapter = new SqlDataAdapter(command);
                dataAdapter.Fill(detallePago);

                int movimientos = detalleFactura.Rows.Count;
                int pagos = detallePago.Rows.Count;
                string seriePago = "";
                int folioPago = 0;
                DateTime fechaPago = DateTime.Now;


                if(pagos == 1)
                {
                    DataRow rowPago = detallePago.Rows[0];
                    seriePago = rowPago["CSERIEDOCUMENTO"].ToString();
                    folioPago = Convert.ToInt32(rowPago["CFOLIO"]);
                    fechaPago = Convert.ToDateTime(rowPago["CFECHA"]);
                    
                }
                else if (pagos > 1)
                {
                    seriePago = "VARIOS";
                    folioPago = 0;
                    fechaPago = DateTime.Now;
                }


                double totalFactura = 0;
                double totalPago = 0;

                //Resume el total de la factura segun los movimientos
                foreach (DataRow rowDetalle in detalleFactura.Rows)
                {
                    totalFactura = totalFactura + Convert.ToDouble(rowDetalle["CTOTAL"]);
                }
                //Resume el total pagado segun los abonos aplicados
                foreach(DataRow rowDetalle in detallePago.Rows)
                {
                    totalPago = totalPago + Convert.ToDouble(rowDetalle["CIMPORTEABONO"]);
                }
                //Provateo y llenado de la tabla Reporte
                foreach(DataRow rowDetalle in detalleFactura.Rows)
                {
                    double totalMovimiento = Convert.ToDouble(rowDetalle["CTOTAL"]);
                    double porcentajePago = (totalMovimiento * 100) / totalPago;
                    porcentajePago = (porcentajePago / 100);
                    double importePagado = (totalPago * porcentajePago);

                    DataRow dataRow = reporte.NewRow();

                    dataRow["CSERIEDOCUMENTO"] = rowDetalle["CSERIEDOCUMENTO"];
                    dataRow["CFOLIO"] = rowDetalle["CFOLIO"];
                    dataRow["CFECHA"] = rowDetalle["CFECHA"];
                    dataRow["CCODIGOCLIENTE"] = rowDetalle["CCODIGOCLIENTE"];
                    dataRow["CRAZONSOCIAL"] = rowDetalle["CRAZONSOCIAL"];
                    dataRow["CRFC"] = rowDetalle["CRFC"];
                    dataRow["CTEXTOEXTRA1"] = rowDetalle["CTEXTOEXTRA1"];
                    dataRow["CTEXTOEXTRA2"] = rowDetalle["CTEXTOEXTRA2"];
                    dataRow["CTEXTOEXTRA3"] = rowDetalle["CTEXTOEXTRA3"];
                    dataRow["CPAIS"] = rowDetalle["CPAIS"];
                    dataRow["CESTADO"] = rowDetalle["CESTADO"];
                    dataRow["CCIUDAD"] = rowDetalle["CCIUDAD"];
                    dataRow["CMUNICIPIO"] = rowDetalle["CMUNICIPIO"];
                    dataRow["CCODIGOPOSTAL"] = rowDetalle["CCODIGOPOSTAL"];
                    dataRow["CCOLONIA"] = rowDetalle["CCOLONIA"];
                    dataRow["CNUMEROINTERIOR"] = rowDetalle["CNUMEROINTERIOR"];
                    dataRow["CNUMEROEXTERIOR"] = rowDetalle["CNUMEROEXTERIOR"];
                    dataRow["CNOMBRECALLE"] = rowDetalle["CNOMBRECALLE"];
                    dataRow["CTELEFONO1"] = rowDetalle["CTELEFONO1"];
                    dataRow["CEMAIL"] = rowDetalle["CEMAIL"];
                    dataRow["TIPOCLIENTE"] = rowDetalle["TIPOCLIENTE"];
                    dataRow["ZONA"] = rowDetalle["ZONA"];
                    dataRow["CCODIGOAGENTE"] = rowDetalle["CCODIGOAGENTE"];
                    dataRow["CNOMBREAGENTE"] = rowDetalle["CNOMBREAGENTE"];
                    dataRow["CCODIGOPRODUCTO"] = rowDetalle["CCODIGOPRODUCTO"];
                    dataRow["CNOMBREPRODUCTO"] = rowDetalle["CNOMBREPRODUCTO"];
                    dataRow["FAMILIA"] = rowDetalle["FAMILIA"];
                    dataRow["SISTEMA"] = rowDetalle["SISTEMA"];
                    dataRow["TIPO"] = rowDetalle["TIPO"];
                    dataRow["CNUMEROMOVIMIENTO"] = rowDetalle["CNUMEROMOVIMIENTO"];
                    dataRow["CUNIDADES"] = rowDetalle["CUNIDADES"];
                    dataRow["CPRECIO"] = rowDetalle["CPRECIO"];
                    dataRow["CNETO"] = rowDetalle["CNETO"];
                    dataRow["CDESCUENTO1"] = rowDetalle["CDESCUENTO1"];
                    dataRow["CDESCUENTO2"] = rowDetalle["CDESCUENTO2"];
                    dataRow["CTOTAL"] = rowDetalle["CTOTAL"];
                    dataRow["CREFERENCIA"] = rowDetalle["CREFERENCIA"];
                    dataRow["COBSERVAMOV"] = rowDetalle["COBSERVAMOV"];
                    dataRow["SERIEPAGO"] = seriePago;
                    dataRow["FOLIOPAGO"] = folioPago;
                    dataRow["FECHAPAGO"] = fechaPago;
                    dataRow["IMPORTEPAGADO"] = importePagado;

                    reporte.Rows.Add(dataRow);

                }
            }


            XLWorkbook workBook = new XLWorkbook();
            workBook.Worksheets.Add(reporte, "Reporte de ventas");

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Hoja de calculo Excel (*.xlsx)|*.xlsx";
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workBook.SaveAs(saveFileDialog.FileName);
                MessageBox.Show("Reporte guardado");
            }
            else
            {
                MessageBox.Show("Error generando el reporte");
            }

        }
    }
}
