using Excel = Microsoft.Office.Interop.Excel;
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
using System.IO;

namespace ReporteComercial.Forms
{
    public partial class ReporteVentas : Form
    {
        string rutaLocal;
        SqlConnection conexion;
        DataTable reporte;

        public ReporteVentas()
        {
            InitializeComponent();
            ConectaBaseDatos();
            CargarDocumentosModelo();
        }

        public void ConectaBaseDatos()
        {
            this.rutaLocal = Application.StartupPath;
            InnerDatabaseManager databaseManager = new InnerDatabaseManager();
            this.conexion = databaseManager.createConnectionFromIniFile(this.rutaLocal);
            //this.conexion.Open();
        }

        public void CargarDocumentosModelo()
        {
            SqlCommand comandoModelo = new SqlCommand("SELECT admDocumentosModelo.CIDDOCUMENTODE, admDocumentosModelo.CDESCRIPCION FROM admDocumentosModelo", this.conexion);
            comandoModelo.CommandType = CommandType.Text;
            SqlDataAdapter dataAdapterModelo = new SqlDataAdapter(comandoModelo);
            DataTable tablaCargo = new DataTable();
            DataTable tablaAbono = new DataTable();
            dataAdapterModelo.Fill(tablaCargo);
            dataAdapterModelo.Fill(tablaAbono);

            this.cmbDoctoCargo.DataSource = tablaCargo;
            this.cmbDoctoCargo.DisplayMember = "CDESCRIPCION";
            this.cmbDoctoCargo.ValueMember = "CIDDOCUMENTODE";

            this.cmbDoctoAbono.DataSource = tablaAbono;
            this.cmbDoctoAbono.DisplayMember = "CDESCRIPCION";
            this.cmbDoctoAbono.ValueMember = "CIDDOCUMENTODE";
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

            int lDoctoCargo = 0;
            int lDoctoAbono = 0;

            lFechaInicial = this.dtpFechaInicial.Text;
            lFechaFinal = this.dtpFechaFinal.Text;

            lDoctoCargo = Convert.ToInt32(this.cmbDoctoCargo.SelectedValue);
            lDoctoAbono = Convert.ToInt32(this.cmbDoctoAbono.SelectedValue);

            Reporte(lDoctoCargo, lDoctoAbono, lFechaInicial, lFechaFinal);
        }

        public void Reporte(int doctoCargo, int doctoAbono, string fechaInicial, string fechaFinal)
        {
            this.reporte = new DataTable();
            this.reporte.Columns.Add("CSERIEDOCUMENTO");
            this.reporte.Columns.Add("CFOLIO");
            this.reporte.Columns.Add("CFECHA");
            this.reporte.Columns.Add("CCODIGOCLIENTE");
            this.reporte.Columns.Add("CRAZONSOCIAL");
            this.reporte.Columns.Add("CRFC");
            this.reporte.Columns.Add("CTEXTOEXTRA1");
            this.reporte.Columns.Add("CTEXTOEXTRA2");
            this.reporte.Columns.Add("CTEXTOEXTRA3");
            this.reporte.Columns.Add("CPAIS");
            this.reporte.Columns.Add("CESTADO");
            this.reporte.Columns.Add("CCIUDAD");
            this.reporte.Columns.Add("CMUNICIPIO");
            this.reporte.Columns.Add("CCODIGOPOSTAL");
            this.reporte.Columns.Add("CCOLONIA");
            this.reporte.Columns.Add("CNUMEROINTERIOR");
            this.reporte.Columns.Add("CNUMEROEXTERIOR");
            this.reporte.Columns.Add("CNOMBRECALLE");
            this.reporte.Columns.Add("CTELEFONO1");
            this.reporte.Columns.Add("CEMAIL");
            this.reporte.Columns.Add("TIPOCLIENTE");
            this.reporte.Columns.Add("ZONA");
            this.reporte.Columns.Add("CCODIGOAGENTE");
            this.reporte.Columns.Add("CNOMBREAGENTE");
            this.reporte.Columns.Add("CCODIGOPRODUCTO");
            this.reporte.Columns.Add("CNOMBREPRODUCTO");
            this.reporte.Columns.Add("FAMILIA");
            this.reporte.Columns.Add("SISTEMA");
            this.reporte.Columns.Add("TIPO");
            this.reporte.Columns.Add("CNUMEROMOVIMIENTO");
            this.reporte.Columns.Add("CUNIDADES");
            this.reporte.Columns.Add("CPRECIO");
            this.reporte.Columns.Add("CNETO");
            this.reporte.Columns.Add("CDESCUENTO1");
            this.reporte.Columns.Add("CDESCUENTO2");
            this.reporte.Columns.Add("SUBTOTAL");
            this.reporte.Columns.Add("CTOTAL");
            this.reporte.Columns.Add("CCOSTOESPECIFICO");
            this.reporte.Columns.Add("UTILIDAD");
            this.reporte.Columns.Add("PORCENTAJEUTILIDAD");
            this.reporte.Columns.Add("CREFERENCIA");
            this.reporte.Columns.Add("COBSERVAMOV");
            this.reporte.Columns.Add("SERIEPAGO");
            this.reporte.Columns.Add("FOLIOPAGO");
            this.reporte.Columns.Add("FECHAPAGO");
            this.reporte.Columns.Add("IMPORTEPAGADO");
            this.reporte.Columns.Add("SALDO");

            //Se modifica el CIDDOCUMENTODE, de (4 facturas) a (19 compras)
            string cmdFacturas = "SELECT CIDDOCUMENTO, CSERIEDOCUMENTO, CFOLIO, CFECHA " +
                "FROM admDocumentos " +
                "WHERE(admDocumentos.CIDDOCUMENTODE = @doctoCargo) " +
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
                "admMovimientos.CTOTAL, admMovimientos.CCOSTOESPECIFICO, " +
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
                "WHERE(admDocumentos.CIDDOCUMENTODE = @doctoCargo AND admDocumentos.CCANCELADO = 0) AND (admDocumentos.CIDDOCUMENTO = @cidfactura) AND (admDomicilios.CTIPOCATALOGO = 1) AND(admDomicilios.CTIPODIRECCION = 0) " +
                "ORDER BY admDocumentos.CFOLIO ASC";

            //Se cambio el CIDDOCUMENTODE de (9 pago del ciente) a (26 abono a proveedor)
            string cmdCargosAbonos = "SELECT        admAsocCargosAbonos.CIDDOCUMENTOABONO, admAsocCargosAbonos.CIDDOCUMENTOCARGO, " +
                "admAsocCargosAbonos.CIMPORTEABONO, admAsocCargosAbonos.CIMPORTECARGO, " +
                "admDocumentos.CSERIEDOCUMENTO, admDocumentos.CFOLIO, admDocumentos.CFECHA, admDocumentos.CTOTAL " +
                "FROM admAsocCargosAbonos " +
                "LEFT JOIN admDocumentos " +
                "ON admAsocCargosAbonos.CIDDOCUMENTOABONO = admDocumentos.CIDDOCUMENTO " +
                "WHERE(admDocumentos.CIDDOCUMENTODE = @doctoAbono) AND (admAsocCargosAbonos.CIDDOCUMENTOCARGO = @cidfactura) AND (admDocumentos.CFECHA BETWEEN @fechaInicial AND @fechaFinal) AND (admDocumentos.CCANCELADO = 0)";

            SqlCommand command = new SqlCommand();
            command = new SqlCommand(cmdFacturas, this.conexion);

            command.Parameters.AddWithValue("@fechaInicial", fechaInicial);
            command.Parameters.AddWithValue("@fechaFinal", fechaFinal);
            command.Parameters.AddWithValue("@doctoCargo", doctoCargo);
            
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
                command.Parameters.AddWithValue("@doctoCargo", doctoCargo);
                dataAdapter = new SqlDataAdapter(command);
                dataAdapter.Fill(detalleFactura);

                //Consulta pagos aplicados
                detallePago.Clear();
                command = new SqlCommand(cmdCargosAbonos, this.conexion);
                command.Parameters.AddWithValue("@cidfactura", cidfactura);
                command.Parameters.AddWithValue("@fechaInicial", fechaInicial);
                command.Parameters.AddWithValue("@fechaFinal", fechaFinal);
                command.Parameters.AddWithValue("@doctoAbono", doctoAbono);
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
                    DataRow rowPago = detallePago.Rows[(pagos-1)];
                    seriePago = "VARIOS";
                    //folioPago = 0;
                    folioPago = Convert.ToInt32(rowPago["CFOLIO"]);
                    //fechaPago = DateTime.Now;
                    fechaPago = Convert.ToDateTime(rowPago["CFECHA"]);
                }


                double totalFactura = 0;
                double totalPago = 0;
                double importePagado = 0;

                //Resume el total de la factura segun los movimientos
                foreach (DataRow rowDetalle in detalleFactura.Rows)
                {
                    totalFactura = totalFactura + Convert.ToDouble(rowDetalle["CTOTAL"] is DBNull ? 0 : rowDetalle["CTOTAL"]);
                }
                //Resume el total pagado segun los abonos aplicados
                foreach(DataRow rowDetalle in detallePago.Rows)
                {
                    totalPago = totalPago + Convert.ToDouble(rowDetalle["CIMPORTEABONO"]);
                }
                //Provateo y llenado de la tabla Reporte
                foreach(DataRow rowDetalle in detalleFactura.Rows)
                {
                    /*
                    if (rowDetalle["CCODIGOCLIENTE"].ToString() == "CLI1421")
                    {
                        MessageBox.Show(totalPago.ToString());
                    }
                    double totalMovimiento = Convert.ToDouble(rowDetalle["CTOTAL"]);
                    double porcentajePago = (totalMovimiento * 100) / totalPago;
                    porcentajePago = (porcentajePago / 100);
                    importePagado = (totalPago * porcentajePago);
                    totalPago = totalPago - importePagado;
                    
                    if (rowDetalle["CCODIGOCLIENTE"].ToString() == "CLI1421")
                    {
                        MessageBox.Show(totalPago.ToString());
                        MessageBox.Show(totalMovimiento.ToString());
                    }
                    */

                    double totalMovimiento = Convert.ToDouble(rowDetalle["CTOTAL"] is DBNull ? 0 : rowDetalle["CTOTAL"]);
                    if (totalPago >= totalMovimiento)
                    {
                        importePagado = totalMovimiento;
                        totalPago = totalPago - importePagado;
                    }
                    else
                    {
                        importePagado = totalPago;
                        totalPago = totalPago - importePagado;
                    }


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

                    double subtotal = (Convert.ToDouble(rowDetalle["CNETO"]) - (Convert.ToDouble(rowDetalle["CDESCUENTO1"]) + Convert.ToDouble(rowDetalle["CDESCUENTO2"])));
                    dataRow["SUBTOTAL"] = subtotal;
                    dataRow["CTOTAL"] = rowDetalle["CTOTAL"];
                    dataRow["CCOSTOESPECIFICO"] = rowDetalle["CCOSTOESPECIFICO"];

                    double utilidad = (subtotal - Convert.ToDouble(rowDetalle["CCOSTOESPECIFICO"]));
                    dataRow["UTILIDAD"] = utilidad;
                    dataRow["PORCENTAJEUTILIDAD"] = (utilidad / subtotal);


                    dataRow["CREFERENCIA"] = rowDetalle["CREFERENCIA"];
                    dataRow["COBSERVAMOV"] = rowDetalle["COBSERVAMOV"];
                    dataRow["SERIEPAGO"] = seriePago;
                    dataRow["FOLIOPAGO"] = folioPago;
                    dataRow["FECHAPAGO"] = fechaPago;
                    dataRow["IMPORTEPAGADO"] = importePagado;
                    dataRow["SALDO"] = (Convert.ToDouble(rowDetalle["CTOTAL"]) - importePagado);

                    reporte.Rows.Add(dataRow);
                }
            }


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Hoja de calculo Excel (*.xlsx)|*.xlsx";
            

            if(saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //workBook.SaveAs(saveFileDialog.FileName);
                ExportarExcel(saveFileDialog.FileName);
                MessageBox.Show("Reporte guardado");
            }
            else
            {
                MessageBox.Show("Error generando el reporte");
            }
        }

        public void ExportarExcel(string fileName)
        {
            Excel.Application excelApp = default(Excel.Application);
            Excel.Workbook workBook = default(Excel.Workbook);
            Excel.Worksheet workSheet = default(Excel.Worksheet);
            excelApp = new Excel.Application();

            workBook = excelApp.Workbooks.Add();
            //workBook.Worksheets.Add();
            workSheet = workBook.Sheets[1];
            workSheet.Activate();


            #region Encabezado
            //Serie
            workSheet.Range["A1"].Font.Bold = true;
            workSheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["A1"].Value = "Serie";
            //Folio
            workSheet.Range["B1"].Font.Bold = true;
            workSheet.Range["B1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["B1"].Value = "Folio folio";
            //Fecha
            workSheet.Range["C1"].Font.Bold = true;
            workSheet.Range["C1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["C1"].Value = "Fecha";
            //Dia
            workSheet.Range["D1"].Font.Bold = true;
            workSheet.Range["D1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["D1"].Value = "Dia";
            //Mes
            workSheet.Range["E1"].Font.Bold = true;
            workSheet.Range["E1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["E1"].Value = "Mes";
            //Año
            workSheet.Range["F1"].Font.Bold = true;
            workSheet.Range["F1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["F1"].Value = "Año";
            //Codigo cliente
            workSheet.Range["G1"].Font.Bold = true;
            workSheet.Range["G1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["G1"].Value = "Codigo del cliente";
            //Razon social
            workSheet.Range["H1"].Font.Bold = true;
            workSheet.Range["H1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["H1"].Value = "Razon social";
            //RFC
            workSheet.Range["I1"].Font.Bold = true;
            workSheet.Range["I1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["I1"].Value = "RFC";
            //Texto Extra 1
            workSheet.Range["J1"].Font.Bold = true;
            workSheet.Range["J1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["J1"].Value = "Texto extra 1";
            //Texto Extra 2
            workSheet.Range["K1"].Font.Bold = true;
            workSheet.Range["K1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["K1"].Value = "Texto extra 2";
            //Texto Extra 3
            workSheet.Range["L1"].Font.Bold = true;
            workSheet.Range["L1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["L1"].Value = "Texto extra 3";
            //Pais
            workSheet.Range["M1"].Font.Bold = true;
            workSheet.Range["M1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["M1"].Value = "Pais";
            //Estado
            workSheet.Range["N1"].Font.Bold = true;
            workSheet.Range["N1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["N1"].Value = "Estado";
            //Ciudad
            workSheet.Range["O1"].Font.Bold = true;
            workSheet.Range["O1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["O1"].Value = "Ciudad";
            //Municipio
            workSheet.Range["P1"].Font.Bold = true;
            workSheet.Range["P1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["P1"].Value = "Municipio";
            //Codigo postal
            workSheet.Range["Q1"].Font.Bold = true;
            workSheet.Range["Q1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["Q1"].Value = "Codigo postal";
            //Colonia
            workSheet.Range["R1"].Font.Bold = true;
            workSheet.Range["R1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["R1"].Value = "Colonia";
            //Calle
            workSheet.Range["S1"].Font.Bold = true;
            workSheet.Range["S1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["S1"].Value = "Calle";
            //Numero exterior
            workSheet.Range["T1"].Font.Bold = true;
            workSheet.Range["T1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["T1"].Value = "Numero exterior";
            //Numero interior
            workSheet.Range["U1"].Font.Bold = true;
            workSheet.Range["U1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["U1"].Value = "Numero interior";
            //Telefono
            workSheet.Range["V1"].Font.Bold = true;
            workSheet.Range["V1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["V1"].Value = "Telefono";
            //E-Mail
            workSheet.Range["W1"].Font.Bold = true;
            workSheet.Range["W1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["W1"].Value = "E-Mail";
            //Tipo Cliente
            workSheet.Range["X1"].Font.Bold = true;
            workSheet.Range["X1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["X1"].Value = "Tipo Cliente";
            //Zona
            workSheet.Range["Y1"].Font.Bold = true;
            workSheet.Range["Y1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["Y1"].Value = "Zona";
            //Codigo agente
            workSheet.Range["Z1"].Font.Bold = true;
            workSheet.Range["Z1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["Z1"].Value = "Codigo del agente";
            //Nombre
            workSheet.Range["AA1"].Font.Bold = true;
            workSheet.Range["AA1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AA1"].Value = "Nombre";
            //Codigo del producto
            workSheet.Range["AB1"].Font.Bold = true;
            workSheet.Range["AB1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AB1"].Value = "Codigo del producto";
            //Nombre
            workSheet.Range["AC1"].Font.Bold = true;
            workSheet.Range["AC1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AC1"].Value = "Nombre";
            //Familia
            workSheet.Range["AD1"].Font.Bold = true;
            workSheet.Range["AD1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AD1"].Value = "Familia";
            //Sistema
            workSheet.Range["AE1"].Font.Bold = true;
            workSheet.Range["AE1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AE1"].Value = "Sistema";
            //Tipo
            workSheet.Range["AF1"].Font.Bold = true;
            workSheet.Range["AF1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AF1"].Value = "Tipo";
            //Numero movimiento
            workSheet.Range["AG1"].Font.Bold = true;
            workSheet.Range["AG1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AG1"].Value = "Numero movimiento";
            //Unidades
            workSheet.Range["AH1"].Font.Bold = true;
            workSheet.Range["AH1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AH1"].Value = "Unidades";
            //Precio
            workSheet.Range["AI1"].Font.Bold = true;
            workSheet.Range["AI1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AI1"].Value = "Precio";
            //Neto
            workSheet.Range["AJ1"].Font.Bold = true;
            workSheet.Range["AJ1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AJ1"].Value = "Neto";
            //Descuento 1
            workSheet.Range["AK1"].Font.Bold = true;
            workSheet.Range["AK1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AK1"].Value = "Descuento 1";
            //Descuento 2
            workSheet.Range["AL1"].Font.Bold = true;
            workSheet.Range["AL1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AL1"].Value = "Descuento 2";
            //Subtotal
            workSheet.Range["AM1"].Font.Bold = true;
            workSheet.Range["AM1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AM1"].Value = "Subtotal";
            //Total
            workSheet.Range["AN1"].Font.Bold = true;
            workSheet.Range["AN1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AN1"].Value = "Total";
            //Costo
            workSheet.Range["AO1"].Font.Bold = true;
            workSheet.Range["AO1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AO1"].Value = "Costo";
            //Utilidad
            workSheet.Range["AP1"].Font.Bold = true;
            workSheet.Range["AP1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AP1"].Value = "Utilidad";
            //Porcentaje utilidad
            workSheet.Range["AQ1"].Font.Bold = true;
            workSheet.Range["AQ1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AQ1"].Value = "Porcentaje utilidad";
            //Referencia
            workSheet.Range["AR1"].Font.Bold = true;
            workSheet.Range["AR1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AR1"].Value = "Referencia";
            //Observacion
            workSheet.Range["AS1"].Font.Bold = true;
            workSheet.Range["AS1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AS1"].Value = "Observacion";
            //Serie pago
            workSheet.Range["AT1"].Font.Bold = true;
            workSheet.Range["AT1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AT1"].Value = "Serie pago";
            //Folio pago
            workSheet.Range["AU1"].Font.Bold = true;
            workSheet.Range["AU1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AU1"].Value = "Folio pago";
            //Fecha pago
            workSheet.Range["AV1"].Font.Bold = true;
            workSheet.Range["AV1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AV1"].Value = "Fecha pago";
            //Dia
            workSheet.Range["AW1"].Font.Bold = true;
            workSheet.Range["AW1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AW1"].Value = "Dia";
            //Mes
            workSheet.Range["AX1"].Font.Bold = true;
            workSheet.Range["AX1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AX1"].Value = "Mes";
            //Año
            workSheet.Range["AY1"].Font.Bold = true;
            workSheet.Range["AY1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AY1"].Value = "Año";
            //Importe pagado
            workSheet.Range["AZ1"].Font.Bold = true;
            workSheet.Range["AZ1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["AZ1"].Value = "Importe pagado";
            //Saldo
            workSheet.Range["BA1"].Font.Bold = true;
            workSheet.Range["BA1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.Range["BA1"].Value = "Saldo";



            #endregion

            #region Detalle
            int actualRow = 1;

            foreach (DataRow row in this.reporte.Rows)
            {
                actualRow++;

                //Serie
                workSheet.Range["A" + actualRow].Value = row["CSERIEDOCUMENTO"].ToString();
                //Folio
                workSheet.Range["B" + actualRow].Value = row["CFOLIO"].ToString();
                //Fecha
                workSheet.Range["C" + actualRow].Value = Convert.ToDateTime(row["CFECHA"]).ToString("dd-MM-yyyy");
                //Dia
                workSheet.Range["D" + actualRow].Value = Convert.ToInt32(Convert.ToDateTime(row["CFECHA"]).Day);
                //Mes
                workSheet.Range["E" + actualRow].Value = Convert.ToInt32(Convert.ToDateTime(row["CFECHA"]).Month);
                //Año
                workSheet.Range["F" + actualRow].Value = Convert.ToInt32(Convert.ToDateTime(row["CFECHA"]).Year);
                //Codigo cliente
                workSheet.Range["G" + actualRow].Value = row["CCODIGOCLIENTE"].ToString();
                //Razon social
                workSheet.Range["H" + actualRow].Value = row["CRAZONSOCIAL"].ToString();
                //RFC
                workSheet.Range["I" + actualRow].Value = row["CRFC"].ToString();
                //Texto Extra 1
                workSheet.Range["J" + actualRow].Value = row["CTEXTOEXTRA1"].ToString();
                //Texto Extra 2
                workSheet.Range["K" + actualRow].Value = row["CTEXTOEXTRA2"].ToString();
                //Texto Extra 3
                workSheet.Range["L" + actualRow].Value = row["CTEXTOEXTRA3"].ToString();
                //Pais
                workSheet.Range["M" + actualRow].Value = row["CPAIS"].ToString();
                //Estado
                workSheet.Range["N" + actualRow].Value = row["CESTADO"].ToString();
                //Ciudad
                workSheet.Range["O" + actualRow].Value = row["CCIUDAD"].ToString();
                //Municipio
                workSheet.Range["P" + actualRow].Value = row["CMUNICIPIO"].ToString();
                //Codigo postal
                workSheet.Range["Q" + actualRow].Value = row["CCODIGOPOSTAL"].ToString();
                //Colonia
                workSheet.Range["R" + actualRow].Value = row["CCOLONIA"].ToString();
                //Calle
                workSheet.Range["S" + actualRow].Value = row["CNUMEROEXTERIOR"].ToString();
                //Numero exterior
                workSheet.Range["T" + actualRow].Value = row["CNUMEROINTERIOR"].ToString();
                //Numero interior
                workSheet.Range["U" + actualRow].Value = row["CNOMBRECALLE"].ToString();
                //Telefono
                workSheet.Range["V" + actualRow].Value = row["CTELEFONO1"].ToString();
                //E-Mail
                workSheet.Range["W" + actualRow].Value = row["CEMAIL"].ToString();
                //Tipo Cliente
                workSheet.Range["X" + actualRow].Value = row["TIPOCLIENTE"].ToString();
                //Zona
                workSheet.Range["Y" + actualRow].Value = row["ZONA"].ToString();
                //Codigo agente
                workSheet.Range["Z" + actualRow].Value = row["CCODIGOAGENTE"].ToString();
                //Nombre
                workSheet.Range["AA" + actualRow].Value = row["CNOMBREAGENTE"].ToString();
                //Codigo del producto
                workSheet.Range["AB" + actualRow].Value = row["CCODIGOPRODUCTO"].ToString();
                //Nombre
                workSheet.Range["AC" + actualRow].Value = row["CNOMBREPRODUCTO"].ToString();
                //Familia
                workSheet.Range["AD" + actualRow].Value = row["FAMILIA"].ToString();
                //Sistema
                workSheet.Range["AE" + actualRow].Value = row["SISTEMA"].ToString();
                //Tipo
                workSheet.Range["AF" + actualRow].Value = row["TIPO"].ToString();
                //Numero movimiento
                workSheet.Range["AG" + actualRow].Value = Convert.ToDouble(row["CNUMEROMOVIMIENTO"]);
                //Unidades
                workSheet.Range["AH" + actualRow].Value = Convert.ToDouble(row["CUNIDADES"]);
                //Precio
                workSheet.Range["AI" + actualRow].Value = Convert.ToDouble(row["CPRECIO"]);
                //Neto
                workSheet.Range["AJ" + actualRow].Value = Convert.ToDouble(row["CNETO"]);
                //Descuento 1
                workSheet.Range["AK" + actualRow].Value = Convert.ToDouble(row["CDESCUENTO1"]);
                //Descuento 2
                workSheet.Range["AL" + actualRow].Value = Convert.ToDouble(row["CDESCUENTO2"]);
                //Subtotal
                workSheet.Range["AM" + actualRow].Value = Convert.ToDouble(row["SUBTOTAL"]);
                //Total
                workSheet.Range["AN" + actualRow].Value = Convert.ToDouble(row["CTOTAL"]);
                //Costo
                workSheet.Range["AO" + actualRow].Value = Convert.ToDouble(row["CCOSTOESPECIFICO"]);
                //Utilidad
                workSheet.Range["AP" + actualRow].Value = Convert.ToDouble(row["UTILIDAD"]);
                //Porcentaje utilidad
                workSheet.Range["AQ" + actualRow].Value = Convert.ToDouble(row["PORCENTAJEUTILIDAD"]);
                //Referencia
                workSheet.Range["AR" + actualRow].Value = row["CREFERENCIA"].ToString();
                //Observacion
                workSheet.Range["AS" + actualRow].Value = row["COBSERVAMOV"].ToString();
                //Serie pago
                workSheet.Range["AT" + actualRow].Value = row["SERIEPAGO"].ToString();
                //Folio pago
                workSheet.Range["AU" + actualRow].Value = row["FOLIOPAGO"].ToString();
                //Fecha pago
                workSheet.Range["AV" + actualRow].Value = Convert.ToDateTime(row["FECHAPAGO"]).ToString("dd-MM-yyyy");
                //Dia
                workSheet.Range["AW" + actualRow].Value = Convert.ToInt32(Convert.ToDateTime(row["FECHAPAGO"]).Day);
                //Mes
                workSheet.Range["AX" + actualRow].Value = Convert.ToInt32(Convert.ToDateTime(row["FECHAPAGO"]).Month);
                //Año
                workSheet.Range["AY" + actualRow].Value = Convert.ToInt32(Convert.ToDateTime(row["FECHAPAGO"]).Year);
                //Importe pagado
                workSheet.Range["AZ" + actualRow].Value = Convert.ToDouble(row["IMPORTEPAGADO"]);
                //Saldo
                workSheet.Range["BA" + actualRow].Value = Convert.ToDouble(row["SALDO"]);

            }

            #endregion

            workSheet.Range["A1" + ":BA1"].Interior.Color = Excel.XlRgbColor.rgbLightGreen;

            workSheet.Range["C2" + ":C" + actualRow].NumberFormat = "yyyy-mm-dd";
            workSheet.Range["AV2" + ":AV" + actualRow].NumberFormat = "yyyy-mm-dd";

            workSheet.Range["AG2" + ":AQ" + actualRow].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            workSheet.Range["AZ2" + ":AZ" + actualRow].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            workSheet.Columns.AutoFit();


            excelApp.Visible = true;
            workSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            workSheet.SaveAs(fileName);
            
        }
    }
}
