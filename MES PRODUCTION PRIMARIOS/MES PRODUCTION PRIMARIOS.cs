using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LaserCuttingApp
{
    public partial class FormLaserCutting : Form
    {
        // ============== CONEXIONES A BASES DE DATOS ==============
        private string connectionStringMartinRea = @"Data Source=10.241.1.27;Initial Catalog=MartinRea MJ;User Id=sa;Password=Mrea$ql123;Connection Timeout=30;Max Pool Size=100;";

        // ============== VARIABLES PRINCIPALES ==============
        private string cncSeleccionado = "";
        private string nestingSeleccionado = "";
        private string parteSeleccionada = "";
        private string recursoSeleccionado = "";
        private string nstRefActual = "";
        private string mnORefActual = "";
        private int cantidadProduccionActual = 0;  // Quantity de la tabla 00000500
        private int cantidadProgramadaNesting = 0;  // Quantity de la tabla 00000100
        private int cantidadReportadaNesting = 0;   // CutQuantity de la tabla 00000100
        private int cantidadPendienteNesting = 0;   // Pendiente del nesting
        private bool isLoading = false;

        // ============== ESTRUCTURAS DE DATOS ==============
        private List<NestingInfo> nestingsActuales = new List<NestingInfo>();
        private List<ParteInfo> partesActuales = new List<ParteInfo>();

        // ============== CONTROLES DE LA INTERFAZ ==============
        private Panel panelCabecera;
        private Panel panelPrincipal;
        private Panel panelInferior;
        private Label lblTitulo;
        private Label lblCNCSeleccionado;
        private Label lblNestingSeleccionado;
        private Label lblRecursoSeleccionado;
        private Label lblParteSeleccionada;
        private Label lblNstRefValor;
        private Label lblCantidadInfo;
        private Label lblEstadoConexion;
        private Label lblFechaActual;
        private Label lblHoraActual;
        private TextBox txtBuscarCNC;
        private FlowLayoutPanel panelNestings;
        private FlowLayoutPanel panelPartes;
        private Button btnBuscar;
        private Button btnReportar;
        private Button btnActualizar;
        private Button btnLimpiar;
        private Timer timerReloj;
        private Label lblCargando;

        // Variables para selección visual
        private Button nestingSeleccionadoBtn = null;
        private Button parteSeleccionadaBtn = null;

        // ============== CONTROLES PARA CANTIDAD ==============
        private TextBox txtCantidad;
        private Label lblCantidadTitulo;

        // ============== CLASES AUXILIARES ==============
        public class NestingInfo
        {
            public string NstRef { get; set; }
            public string Nombre { get; set; }
            public int CantidadProgramada { get; set; }
            public int CantidadReportada { get; set; }
            public int CantidadPendiente { get; set; }
            public string Estado { get; set; }
        }

        public class ParteInfo
        {
            public string MnORef { get; set; }      // JOB
            public string OprID { get; set; }       // Operación
            public string PrdRefDst { get; set; }   // Número de parte
            public string Recurso { get; set; }     // Recurso asociado
            public int Cantidad { get; set; }       // Quantity de la tabla 00000500
            public string NstRef { get; set; }
        }

        // ============== COLORES DEL SISTEMA ==============
        private Color colorPrimario = Color.FromArgb(106, 13, 18);
        private Color colorSecundario = Color.FromArgb(0, 114, 198);
        private Color colorExito = Color.FromArgb(40, 167, 69);
        private Color colorPeligro = Color.FromArgb(220, 53, 69);
        private Color colorInfo = Color.FromArgb(23, 162, 184);
        private Color colorFondo = Color.FromArgb(245, 245, 245);
        private Color colorNormal = Color.FromArgb(248, 249, 250);
        private Color colorSeleccionado = Color.FromArgb(0, 114, 198);
        private Color colorBorde = Color.FromArgb(200, 200, 200);

        public FormLaserCutting()
        {
            InitializeComponent();
            ConfigurarEstiloGeneral();
            ConfigurarInterfaz();
            CargarDatosInicialesAsync();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ClientSize = new System.Drawing.Size(1366, 800);
            this.Name = "FormLaserCutting";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
        }

        private void ConfigurarEstiloGeneral()
        {
            this.Text = "SISTEMA DE REPORTE - CORTADORAS LÁSER";
            this.BackColor = colorFondo;
            this.DoubleBuffered = true;

            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.Size = new Size(1366, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Normal;
        }

        private void ConfigurarInterfaz()
        {
            ConfigurarCabecera();
            ConfigurarPanelPrincipal();
            ConfigurarPanelInferior();
            ConfigurarTimers();
        }

        private void ConfigurarCabecera()
        {
            panelCabecera = new Panel
            {
                Dock = DockStyle.Top,
                Height = 70,
                BackColor = colorPrimario
            };

            lblTitulo = new Label
            {
                Text = "REPORTE DE PRODUCCIÓN - CORTADORAS LÁSER",
                Font = new Font("Arial", 18, FontStyle.Bold),
                ForeColor = Color.White,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };

            panelCabecera.Controls.Add(lblTitulo);
            this.Controls.Add(panelCabecera);
        }

        private void ConfigurarPanelPrincipal()
        {
            panelPrincipal = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(15)
            };

            // ========== PANEL DE BÚSQUEDA ==========
            Panel panelBusqueda = new Panel
            {
                Location = new Point(15, 50),
                Size = new Size(1330, 100),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            Label lblBuscarCNC = new Label
            {
                Text = "BUSCAR CNC:",
                Font = new Font("Arial", 12, FontStyle.Bold),
                ForeColor = colorPrimario,
                Location = new Point(15, 25),
                AutoSize = true
            };

            txtBuscarCNC = new TextBox
            {
                Location = new Point(140, 22),
                Size = new Size(250, 30),
                Font = new Font("Arial", 12),
                Text = "Ingrese código CNC...",
                ForeColor = Color.Gray,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            txtBuscarCNC.GotFocus += TxtBuscarCNC_GotFocus;
            txtBuscarCNC.LostFocus += TxtBuscarCNC_LostFocus;

            btnBuscar = new Button
            {
                Text = "BUSCAR",
                Location = new Point(395, 20),
                Size = new Size(100, 35),
                Font = new Font("Arial", 10, FontStyle.Bold),
                BackColor = colorInfo,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnBuscar.FlatAppearance.BorderSize = 0;
            btnBuscar.Click += async (s, e) => await BuscarCNCAsync();

            Button btnRefrescar = new Button
            {
                Text = "REFRESCAR",
                Location = new Point(510, 20),
                Size = new Size(100, 35),
                Font = new Font("Arial", 10, FontStyle.Bold),
                BackColor = colorExito,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnRefrescar.FlatAppearance.BorderSize = 0;
            btnRefrescar.Click += async (s, e) => await RefrescarDatosAsync();

            Label lblEjemplo = new Label
            {
                Text = "Ejemplo: 003287, 001234, 005678",
                Font = new Font("Arial", 9),
                ForeColor = Color.Gray,
                Location = new Point(130, 60),
                AutoSize = true
            };

            panelBusqueda.Controls.Add(lblBuscarCNC);
            panelBusqueda.Controls.Add(txtBuscarCNC);
            panelBusqueda.Controls.Add(btnBuscar);
            panelBusqueda.Controls.Add(btnRefrescar);
            panelBusqueda.Controls.Add(lblEjemplo);

            // ========== LABEL DE CARGA ==========
            lblCargando = new Label
            {
                Text = "Cargando datos...",
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorInfo,
                Location = new Point(15, 125),
                AutoSize = true,
                Visible = false
            };
            panelPrincipal.Controls.Add(lblCargando);

            // ========== PANEL DE NESTINGS ==========
            Panel panelNestingsContainer = new Panel
            {
                Location = new Point(15, 160),
                Size = new Size(430, 520),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            Label lblNestingsTitulo = new Label
            {
                Text = "NESTINGS / TRABAJOS DEL CNC",
                Font = new Font("Arial", 12, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(10, 10),
                AutoSize = true
            };

            Label lblInfoNestings = new Label
            {
                Text = "Seleccione un nesting para ver sus partes",
                Font = new Font("Arial", 9),
                ForeColor = Color.Gray,
                Location = new Point(10, 35),
                AutoSize = true
            };

            panelNestings = new FlowLayoutPanel
            {
                Location = new Point(10, 60),
                Size = new Size(405, 445),
                AutoScroll = true,
                Padding = new Padding(5),
                WrapContents = true,
                BackColor = Color.White
            };

            panelNestingsContainer.Controls.Add(lblNestingsTitulo);
            panelNestingsContainer.Controls.Add(lblInfoNestings);
            panelNestingsContainer.Controls.Add(panelNestings);

            // ========== PANEL DE PARTES ==========
            Panel panelPartesContainer = new Panel
            {
                Location = new Point(460, 160),
                Size = new Size(430, 520),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            Label lblPartesTitulo = new Label
            {
                Text = "NÚMEROS DE PARTE",
                Font = new Font("Arial", 12, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(10, 10),
                AutoSize = true
            };

            Label lblInfoPartes = new Label
            {
                Text = "Seleccione una parte para reportar",
                Font = new Font("Arial", 9),
                ForeColor = Color.Gray,
                Location = new Point(10, 35),
                AutoSize = true
            };

            panelPartes = new FlowLayoutPanel
            {
                Location = new Point(10, 60),
                Size = new Size(405, 445),
                AutoScroll = true,
                Padding = new Padding(5),
                WrapContents = true,
                BackColor = Color.White
            };

            panelPartesContainer.Controls.Add(lblPartesTitulo);
            panelPartesContainer.Controls.Add(lblInfoPartes);
            panelPartesContainer.Controls.Add(panelPartes);

            // ========== PANEL DE CONTROL ==========
            Panel panelControl = new Panel
            {
                Location = new Point(905, 160),
                Size = new Size(440, 520),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            Label lblInfoTrabajo = new Label
            {
                Text = "INFORMACIÓN DEL TRABAJO",
                Font = new Font("Arial", 12, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(15, 10),
                AutoSize = true
            };

            Panel panelInfoTrabajo = new Panel
            {
                Location = new Point(15, 40),
                Size = new Size(410, 150),
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };

            lblCNCSeleccionado = new Label
            {
                Text = "CNC: --",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 10),
                AutoSize = true
            };

            lblNestingSeleccionado = new Label
            {
                Text = "Nesting: --",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 35),
                AutoSize = true
            };

            lblRecursoSeleccionado = new Label
            {
                Text = "Recurso: --",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 60),
                AutoSize = true
            };

            lblParteSeleccionada = new Label
            {
                Text = "Parte: --",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 85),
                AutoSize = true
            };

            Label lblJobInfo = new Label
            {
                Text = "JOB / Operación:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 110),
                AutoSize = true
            };

            Label lblJobValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = colorInfo,
                Location = new Point(130, 110),
                AutoSize = true,
                Name = "lblJobValor"
            };

            Label lblNstRef = new Label
            {
                Text = "NstRef:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 135),
                AutoSize = true
            };

            lblNstRefValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = colorInfo,
                Location = new Point(70, 135),
                AutoSize = true
            };

            panelInfoTrabajo.Controls.Add(lblCNCSeleccionado);
            panelInfoTrabajo.Controls.Add(lblNestingSeleccionado);
            panelInfoTrabajo.Controls.Add(lblRecursoSeleccionado);
            panelInfoTrabajo.Controls.Add(lblParteSeleccionada);
            panelInfoTrabajo.Controls.Add(lblJobInfo);
            panelInfoTrabajo.Controls.Add(lblJobValor);
            panelInfoTrabajo.Controls.Add(lblNstRef);
            panelInfoTrabajo.Controls.Add(lblNstRefValor);

            lblCantidadInfo = new Label
            {
                Text = "Nesting - Programado: 0 | Reportado: 0 | Pendiente: 0",
                Font = new Font("Arial", 9),
                ForeColor = colorPrimario,
                Location = new Point(15, 200),
                AutoSize = true
            };

            Label lblSeparador = new Label
            {
                Text = "────────────────────────────────────────────",
                Font = new Font("Arial", 8),
                ForeColor = Color.Gray,
                Location = new Point(15, 225),
                AutoSize = true
            };

            lblCantidadTitulo = new Label
            {
                Text = "CANTIDAD A REPORTAR (Quantity de la parte)",
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(15, 245),
                AutoSize = true
            };

            txtCantidad = new TextBox
            {
                Location = new Point(15, 275),
                Size = new Size(410, 45),
                Font = new Font("Arial", 18, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
                Text = "0",
                BorderStyle = BorderStyle.FixedSingle,
                ReadOnly = true
            };

            btnReportar = new Button
            {
                Text = "REPORTAR PRODUCCIÓN",
                Location = new Point(15, 340),
                Size = new Size(410, 50),
                Font = new Font("Arial", 11, FontStyle.Bold),
                BackColor = colorExito,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnReportar.FlatAppearance.BorderSize = 0;
            btnReportar.Click += async (s, e) => await ReportarProduccionAsync();
            btnReportar.Enabled = false;

            btnActualizar = new Button
            {
                Text = "ACTUALIZAR CANTIDAD PROGRAMADA DEL NESTING",
                Location = new Point(15, 400),
                Size = new Size(410, 50),
                Font = new Font("Arial", 10, FontStyle.Bold),
                BackColor = colorInfo,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnActualizar.FlatAppearance.BorderSize = 0;
            btnActualizar.Click += async (s, e) => await ActualizarCantidadProgramadaNestingAsync();
            btnActualizar.Enabled = false;

            btnLimpiar = new Button
            {
                Text = "LIMPIAR SELECCIÓN",
                Location = new Point(15, 460),
                Size = new Size(410, 45),
                Font = new Font("Arial", 11, FontStyle.Bold),
                BackColor = colorPeligro,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnLimpiar.FlatAppearance.BorderSize = 0;
            btnLimpiar.Click += (s, e) => LimpiarSeleccion();

            panelControl.Controls.Add(lblInfoTrabajo);
            panelControl.Controls.Add(panelInfoTrabajo);
            panelControl.Controls.Add(lblCantidadInfo);
            panelControl.Controls.Add(lblSeparador);
            panelControl.Controls.Add(lblCantidadTitulo);
            panelControl.Controls.Add(txtCantidad);
            panelControl.Controls.Add(btnReportar);
            panelControl.Controls.Add(btnActualizar);
            panelControl.Controls.Add(btnLimpiar);

            panelPrincipal.Controls.Add(panelBusqueda);
            panelPrincipal.Controls.Add(panelNestingsContainer);
            panelPrincipal.Controls.Add(panelPartesContainer);
            panelPrincipal.Controls.Add(panelControl);

            this.Controls.Add(panelPrincipal);
        }

         private void ConfigurarPanelInferior()
        {
            panelInferior = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 45,
                BackColor = Color.FromArgb(240, 240, 240),
                BorderStyle = BorderStyle.FixedSingle
            };

            lblEstadoConexion = new Label
            {
                Text = "● CONECTADO",
                Font = new Font("Arial", 9, FontStyle.Bold),
                ForeColor = Color.Green,
                Location = new Point(15, 12),
                AutoSize = true
            };

            lblFechaActual = new Label
            {
                Text = DateTime.Now.ToString("dd/MM/yyyy"),
                Font = new Font("Arial", 9, FontStyle.Bold),
                Location = new Point(200, 12),
                AutoSize = true
            };

            lblHoraActual = new Label
            {
                Text = DateTime.Now.ToString("HH:mm:ss"),
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(200, 25),
                AutoSize = true
            };

            Label lblVersion = new Label
            {
                Text = "ACTUALIZADO CADA DÍA EN PROCESAMIENTO",
                Font = new Font("Arial", 8),
                ForeColor = Color.Gray,
                Location = new Point(1150, 15),
                AutoSize = true
            };

            panelInferior.Controls.Add(lblEstadoConexion);
            panelInferior.Controls.Add(lblFechaActual);
            panelInferior.Controls.Add(lblHoraActual);
            panelInferior.Controls.Add(lblVersion);

            this.Controls.Add(panelInferior);
        }

        private void ConfigurarTimers()
        {
            timerReloj = new Timer { Interval = 1000, Enabled = true };
            timerReloj.Tick += (s, e) =>
            {
                lblHoraActual.Text = DateTime.Now.ToString("HH:mm:ss");
                lblFechaActual.Text = DateTime.Now.ToString("dd/MM/yyyy");
            };
        }

        // ============== MÉTODOS PARA PLACEHOLDER ==============
        private void TxtBuscarCNC_GotFocus(object sender, EventArgs e)
        {
            if (txtBuscarCNC.Text == "Ingrese código CNC...")
            {
                txtBuscarCNC.Text = "";
                txtBuscarCNC.ForeColor = Color.Black;
            }
        }

        private void TxtBuscarCNC_LostFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBuscarCNC.Text))
            {
                txtBuscarCNC.Text = "Ingrese código CNC...";
                txtBuscarCNC.ForeColor = Color.Gray;
            }
        }

        // ============== MÉTODOS DE DATOS ==============

        private async Task CargarDatosInicialesAsync()
        {
            try
            {
                await ProbarConexionAsync();
                MostrarNotificacion("Sistema listo. Ingrese un código CNC para buscar.", Color.Green);
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al conectar: {ex.Message}", Color.Red);
                lblEstadoConexion.Text = "● DESCONECTADO";
                lblEstadoConexion.ForeColor = Color.Red;
            }
        }

        private Task ProbarConexionAsync()
        {
            return Task.Run(() =>
            {
                using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                {
                    conn.Open();
                }
            });
        }

        private async Task BuscarCNCAsync()
        {
            string cncABuscar = txtBuscarCNC.Text;
            if (cncABuscar == "Ingrese código CNC..." || string.IsNullOrWhiteSpace(cncABuscar))
            {
                MostrarNotificacion("Ingrese un código CNC válido", Color.Orange);
                return;
            }

            if (isLoading) return;
            isLoading = true;

            MostrarCargando(true, $"Buscando CNC: {cncABuscar}...");

            try
            {
                string queryNestings = @"
                SELECT 
                    NstRef,
                    Name as Nombre,
                    Quantity as Cantidad_Programada,
                    ISNULL(CutQuantity, 0) as Cantidad_Reportada,
                    ISNULL(Quantity, 0) - ISNULL(CutQuantity, 0) as Cantidad_Pendiente,
                    CuttingStatus as Estado
                FROM [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000100]
                WHERE CNC = @cnc
                ORDER BY Name";

                DataTable dtNestings = new DataTable();

                await Task.Run(() =>
                {
                    using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                    {
                        using (SqlCommand cmd = new SqlCommand(queryNestings, conn))
                        {
                            cmd.Parameters.AddWithValue("@cnc", cncABuscar);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dtNestings);
                            }
                        }
                    }
                });

                nestingsActuales.Clear();
                foreach (DataRow row in dtNestings.Rows)
                {
                    nestingsActuales.Add(new NestingInfo
                    {
                        NstRef = row["NstRef"].ToString(),
                        Nombre = row["Nombre"].ToString(),
                        CantidadProgramada = Convert.ToInt32(row["Cantidad_Programada"]),
                        CantidadReportada = Convert.ToInt32(row["Cantidad_Reportada"]),
                        CantidadPendiente = Convert.ToInt32(row["Cantidad_Pendiente"]),
                        Estado = row["Estado"].ToString()
                    });
                }

                cncSeleccionado = cncABuscar;
                lblCNCSeleccionado.Text = $"CNC: {cncSeleccionado}";
                CrearBotonesNestings();

                if (nestingsActuales.Count == 0)
                {
                    MostrarNotificacion($"No se encontraron nestings para el CNC: {cncABuscar}", Color.Orange);
                }
                else
                {
                    MostrarNotificacion($"Se encontraron {nestingsActuales.Count} nestings para el CNC: {cncABuscar}", Color.Green);
                }
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al buscar CNC: {ex.Message}", Color.Red);
            }
            finally
            {
                isLoading = false;
                MostrarCargando(false);
            }
        }

        private async Task CargarPartesPorNestingAsync(string nstRef, string nestingNombre)
        {
            if (isLoading) return;
            isLoading = true;

            MostrarCargando(true, "Cargando partes...");

            try
            {
                // Buscar en DIS_NEST_NEST_00000500 para obtener las partes con su cantidad
                string query = @"
                SELECT 
                    MnORef,
                    OprID,
                    PrdRefDst,
                    Quantity
                FROM [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000500]
                WHERE NstRef = @nstRef
                AND PrdRefDst IS NOT NULL
                AND PrdRefDst != ''
                ORDER BY PrdRefDst";

                DataTable dt = new DataTable();

                await Task.Run(() =>
                {
                    using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                    {
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@nstRef", nstRef);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                });

                partesActuales.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    string mnORef = row["MnORef"].ToString();
                    string oprID = row["OprID"].ToString();
                    string prdRefDst = row["PrdRefDst"].ToString();
                    int cantidad = Convert.ToInt32(row["Quantity"]);

                    partesActuales.Add(new ParteInfo
                    {
                        MnORef = mnORef,
                        OprID = oprID,
                        PrdRefDst = prdRefDst,
                        Recurso = nestingNombre,
                        Cantidad = cantidad,
                        NstRef = nstRef
                    });
                }

                nestingSeleccionado = nestingNombre;
                lblNestingSeleccionado.Text = $"Nesting: {nestingSeleccionado}";
                CrearBotonesPartes();

                if (partesActuales.Count == 0)
                {
                    MostrarNotificacion($"No se encontraron partes para el nesting: {nestingNombre}", Color.Orange);
                }
                else
                {
                    MostrarNotificacion($"Se encontraron {partesActuales.Count} partes para el nesting: {nestingNombre}", Color.Green);
                }
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al cargar Partes: {ex.Message}", Color.Red);
            }
            finally
            {
                isLoading = false;
                MostrarCargando(false);
            }
        }

        private async Task ActualizarInfoParteAsync(ParteInfo parte)
        {
            try
            {
                // Obtener información del nesting para mostrar programado/reportado/pendiente
                string queryNesting = @"
                SELECT 
                    Name as Recurso,
                    Quantity as Cantidad_Programada,
                    ISNULL(CutQuantity, 0) as Cantidad_Reportada,
                    ISNULL(Quantity, 0) - ISNULL(CutQuantity, 0) as Cantidad_Pendiente
                FROM [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000100]
                WHERE NstRef = @nstRef";

                DataTable dtNesting = new DataTable();

                await Task.Run(() =>
                {
                    using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                    {
                        using (SqlCommand cmd = new SqlCommand(queryNesting, conn))
                        {
                            cmd.Parameters.AddWithValue("@nstRef", parte.NstRef);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dtNesting);
                            }
                        }
                    }
                });

                if (dtNesting.Rows.Count > 0)
                {
                    DataRow row = dtNesting.Rows[0];
                    recursoSeleccionado = row["Recurso"].ToString();
                    cantidadProgramadaNesting = Convert.ToInt32(row["Cantidad_Programada"]);
                    cantidadReportadaNesting = Convert.ToInt32(row["Cantidad_Reportada"]);
                    cantidadPendienteNesting = Convert.ToInt32(row["Cantidad_Pendiente"]);

                    nstRefActual = parte.NstRef;
                    mnORefActual = parte.MnORef;
                    parteSeleccionada = parte.PrdRefDst;
                    cantidadProduccionActual = parte.Cantidad;

                    // Actualizar labels
                    lblRecursoSeleccionado.Text = $"Recurso: {recursoSeleccionado}";
                    lblParteSeleccionada.Text = $"Parte: {parteSeleccionada}";

                    // Buscar el label de Job
                    Control[] controls = panelPrincipal.Controls.Find("lblJobValor", true);
                    if (controls.Length > 0)
                    {
                        ((Label)controls[0]).Text = $"{parte.MnORef} / {parte.OprID}";
                    }

                    lblNstRefValor.Text = nstRefActual;
                    lblCantidadInfo.Text = $"Nesting - Programado: {cantidadProgramadaNesting:N0} | Reportado: {cantidadReportadaNesting:N0} | Pendiente: {cantidadPendienteNesting:N0}";

                    // Establecer la cantidad de producción en el TextBox
                    txtCantidad.Text = cantidadProduccionActual.ToString("N0");
                    txtCantidad.ForeColor = cantidadProduccionActual > 0 ? Color.Green : Color.Gray;

                    btnReportar.Enabled = cantidadProduccionActual > 0;
                    btnActualizar.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al obtener información: {ex.Message}", Color.Red);
            }
        }

        // ============== CREACIÓN DE BOTONES ==============

        private void CrearBotonesNestings()
        {
            panelNestings.Controls.Clear();

            foreach (var nesting in nestingsActuales)
            {
                string estadoIcono = nesting.Estado == "Completado" ? "✅" : (nesting.CantidadPendiente > 0 ? "📋" : "⚪");
                string texto = $"{estadoIcono} {nesting.Nombre}\nPendiente: {nesting.CantidadPendiente:N0}";

                Button btnNesting = new Button
                {
                    Text = texto,
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    BackColor = Color.White,
                    ForeColor = Color.Black,
                    Size = new Size(385, 55),
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand,
                    Margin = new Padding(3),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Tag = nesting
                };
                btnNesting.FlatAppearance.BorderSize = 1;
                btnNesting.FlatAppearance.BorderColor = colorBorde;
                btnNesting.Click += BtnNesting_Click;

                panelNestings.Controls.Add(btnNesting);
            }
        }

        private void CrearBotonesPartes()
        {
            panelPartes.Controls.Clear();

            int index = 1;
            foreach (var parte in partesActuales)
            {
                Button btnParte = new Button
                {
                    Text = $"{index:00}. {parte.PrdRefDst}\nCantidad: {parte.Cantidad:N0}",
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    BackColor = Color.White,
                    ForeColor = Color.Black,
                    Size = new Size(385, 55),
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand,
                    Margin = new Padding(3),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Tag = parte
                };
                btnParte.FlatAppearance.BorderSize = 1;
                btnParte.FlatAppearance.BorderColor = colorBorde;
                btnParte.Click += BtnParte_Click;

                panelPartes.Controls.Add(btnParte);
                index++;
            }
        }

        // ============== EVENTOS ==============

        private async void BtnNesting_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            NestingInfo nesting = (NestingInfo)btn.Tag;

            if (nestingSeleccionadoBtn != null)
            {
                nestingSeleccionadoBtn.BackColor = Color.White;
                nestingSeleccionadoBtn.ForeColor = Color.Black;
            }

            btn.BackColor = colorSeleccionado;
            btn.ForeColor = Color.White;
            nestingSeleccionadoBtn = btn;

            if (parteSeleccionadaBtn != null)
            {
                parteSeleccionadaBtn.BackColor = Color.White;
                parteSeleccionadaBtn.ForeColor = Color.Black;
                parteSeleccionadaBtn = null;
            }

            lblParteSeleccionada.Text = "Parte: --";
            lblRecursoSeleccionado.Text = "Recurso: --";
            lblNstRefValor.Text = "--";
            lblCantidadInfo.Text = "Nesting - Programado: 0 | Reportado: 0 | Pendiente: 0";
            parteSeleccionada = "";
            recursoSeleccionado = "";
            nstRefActual = "";
            txtCantidad.Text = "0";
            btnReportar.Enabled = false;
            btnActualizar.Enabled = false;

            await CargarPartesPorNestingAsync(nesting.NstRef, nesting.Nombre);
        }

        private async void BtnParte_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            ParteInfo parte = (ParteInfo)btn.Tag;

            if (parteSeleccionadaBtn != null)
            {
                parteSeleccionadaBtn.BackColor = Color.White;
                parteSeleccionadaBtn.ForeColor = Color.Black;
            }

            btn.BackColor = colorSeleccionado;
            btn.ForeColor = Color.White;
            parteSeleccionadaBtn = btn;

            await ActualizarInfoParteAsync(parte);
        }

        // ============== MÉTODOS DE REPORTE ==============

        private int ObtenerCantidadIngresada()
        {
            if (string.IsNullOrEmpty(txtCantidad.Text))
                return 0;

            string valorLimpio = txtCantidad.Text.Replace(",", "");
            return int.TryParse(valorLimpio, out int resultado) ? resultado : 0;
        }

        private async Task ReportarProduccionAsync()
        {
            if (string.IsNullOrEmpty(nstRefActual))
            {
                MostrarNotificacion("No se encontró información para esta parte", Color.Red);
                return;
            }

            int cantidadReportar = ObtenerCantidadIngresada();

            if (cantidadReportar <= 0)
            {
                MostrarNotificacion("La cantidad debe ser mayor a 0", Color.Red);
                return;
            }

            DialogResult result = MessageBox.Show(
                $"¿Confirmar reporte de {cantidadReportar:N0} piezas?\n\n" +
                $"CNC: {cncSeleccionado}\n" +
                $"Nesting: {nestingSeleccionado}\n" +
                $"Recurso: {recursoSeleccionado}\n" +
                $"Parte: {parteSeleccionada}\n" +
                $"JOB: {mnORefActual}\n" +
                $"Cantidad de producción: {cantidadProduccionActual:N0}",
                "Confirmar Reporte",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    btnReportar.Enabled = false;
                    btnReportar.Text = "PROCESANDO...";

                    await Task.Run(() =>
                    {
                        // Actualizar CutQuantity en la tabla 00000100
                        string queryUpdate = @"
                        UPDATE [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000100]
                        SET 
                            CutQuantity = ISNULL(CutQuantity, 0) + @cantidad,
                            CuttingStatus = CASE 
                                WHEN ISNULL(Quantity, 0) <= (ISNULL(CutQuantity, 0) + @cantidad) THEN 'Completado'
                                ELSE 'En Proceso'
                            END,
                            LastDate = GETDATE(),
                            LastUser = @usuario
                        WHERE NstRef = @nstRef";

                        using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                        {
                            using (SqlCommand cmd = new SqlCommand(queryUpdate, conn))
                            {
                                cmd.Parameters.AddWithValue("@cantidad", cantidadReportar);
                                cmd.Parameters.AddWithValue("@nstRef", nstRefActual);
                                cmd.Parameters.AddWithValue("@usuario", Environment.UserName);
                                conn.Open();
                                cmd.ExecuteNonQuery();
                            }
                        }
                    });

                    MostrarNotificacion($"✓ Reporte exitoso: {cantidadReportar:N0} piezas", Color.Green);
                    await RegistrarLogReporteAsync(cantidadReportar);

                    // Recargar información actualizada
                    await BuscarCNCAsync();

                    if (!string.IsNullOrEmpty(nstRefActual))
                    {
                        var parte = partesActuales.FirstOrDefault(p => p.PrdRefDst == parteSeleccionada);
                        if (parte != null)
                        {
                            await ActualizarInfoParteAsync(parte);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MostrarNotificacion($"Error al reportar: {ex.Message}", Color.Red);
                }
                finally
                {
                    btnReportar.Text = "REPORTAR PRODUCCIÓN";
                    btnReportar.Enabled = true;
                }
            }
        }

        private async Task ActualizarCantidadProgramadaNestingAsync()
        {
            if (string.IsNullOrEmpty(nstRefActual))
            {
                MostrarNotificacion("No se encontró información para este nesting", Color.Red);
                return;
            }

            int nuevaCantidad = ObtenerCantidadIngresada();

            if (nuevaCantidad <= 0)
            {
                MostrarNotificacion("La cantidad debe ser mayor a 0", Color.Red);
                return;
            }

            DialogResult result = MessageBox.Show(
                $"¿Actualizar cantidad programada del nesting a {nuevaCantidad:N0} piezas?\n\n" +
                $"Nesting: {nestingSeleccionado}\n" +
                $"NstRef: {nstRefActual}",
                "Confirmar Actualización",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                try
                {
                    btnActualizar.Enabled = false;
                    btnActualizar.Text = "PROCESANDO...";

                    await Task.Run(() =>
                    {
                        string query = @"
                        UPDATE [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000100]
                        SET 
                            Quantity = @nuevaCantidad,
                            LastDate = GETDATE(),
                            LastUser = @usuario
                        WHERE NstRef = @nstRef";

                        using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                        {
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@nuevaCantidad", nuevaCantidad);
                                cmd.Parameters.AddWithValue("@nstRef", nstRefActual);
                                cmd.Parameters.AddWithValue("@usuario", Environment.UserName);
                                conn.Open();
                                cmd.ExecuteNonQuery();
                            }
                        }
                    });

                    MostrarNotificacion($"✓ Cantidad programada actualizada a {nuevaCantidad:N0}", Color.Green);
                    await BuscarCNCAsync();

                    if (!string.IsNullOrEmpty(nstRefActual))
                    {
                        var parte = partesActuales.FirstOrDefault(p => p.PrdRefDst == parteSeleccionada);
                        if (parte != null)
                        {
                            await ActualizarInfoParteAsync(parte);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MostrarNotificacion($"Error al actualizar: {ex.Message}", Color.Red);
                }
                finally
                {
                    btnActualizar.Text = "ACTUALIZAR CANTIDAD PROGRAMADA DEL NESTING";
                    btnActualizar.Enabled = true;
                }
            }
        }

        private async Task RegistrarLogReporteAsync(int cantidad)
        {
            await Task.Run(() =>
            {
                try
                {
                    string createTableQuery = @"
                    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ReportesProduccionLaser' AND xtype='U')
                    CREATE TABLE ReportesProduccionLaser (
                        ID INT IDENTITY(1,1) PRIMARY KEY,
                        CNC NVARCHAR(50),
                        NstRef NVARCHAR(100),
                        Nesting NVARCHAR(200),
                        Recurso NVARCHAR(200),
                        Parte NVARCHAR(200),
                        MnORef NVARCHAR(100),
                        OprID NVARCHAR(50),
                        CantidadReportada INT,
                        CantidadProduccion INT,
                        FechaReporte DATETIME,
                        Usuario NVARCHAR(100),
                        Observaciones NVARCHAR(500)
                    )";

                    using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                    {
                        conn.Open();
                        using (SqlCommand cmdCreate = new SqlCommand(createTableQuery, conn))
                        {
                            cmdCreate.ExecuteNonQuery();
                        }
                    }

                    string query = @"
                    INSERT INTO ReportesProduccionLaser 
                    (CNC, NstRef, Nesting, Recurso, Parte, MnORef, OprID, CantidadReportada, CantidadProduccion, FechaReporte, Usuario, Observaciones)
                    VALUES (@cnc, @nstRef, @nesting, @recurso, @parte, @mnORef, @oprID, @cantidad, @cantidadProduccion, GETDATE(), @usuario, @observaciones)";

                    using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                    {
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@cnc", cncSeleccionado);
                            cmd.Parameters.AddWithValue("@nstRef", nstRefActual);
                            cmd.Parameters.AddWithValue("@nesting", nestingSeleccionado);
                            cmd.Parameters.AddWithValue("@recurso", recursoSeleccionado);
                            cmd.Parameters.AddWithValue("@parte", parteSeleccionada);
                            cmd.Parameters.AddWithValue("@mnORef", mnORefActual);
                            cmd.Parameters.AddWithValue("@oprID", "");
                            cmd.Parameters.AddWithValue("@cantidad", cantidad);
                            cmd.Parameters.AddWithValue("@cantidadProduccion", cantidadProduccionActual);
                            cmd.Parameters.AddWithValue("@usuario", Environment.UserName);
                            cmd.Parameters.AddWithValue("@observaciones", "Reporte desde aplicación láser");
                            conn.Open();
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error en log: {ex.Message}");
                }
            });
        }

        private async Task RefrescarDatosAsync()
        {
            if (!string.IsNullOrEmpty(cncSeleccionado))
            {
                await BuscarCNCAsync();
            }
            else
            {
                MostrarNotificacion("Primero busque un CNC", Color.Orange);
            }
        }

        private void LimpiarSeleccion()
        {
            if (nestingSeleccionadoBtn != null)
            {
                nestingSeleccionadoBtn.BackColor = Color.White;
                nestingSeleccionadoBtn.ForeColor = Color.Black;
                nestingSeleccionadoBtn = null;
            }
            if (parteSeleccionadaBtn != null)
            {
                parteSeleccionadaBtn.BackColor = Color.White;
                parteSeleccionadaBtn.ForeColor = Color.Black;
                parteSeleccionadaBtn = null;
            }

            nestingSeleccionado = "";
            parteSeleccionada = "";
            recursoSeleccionado = "";
            nstRefActual = "";
            mnORefActual = "";
            cantidadProduccionActual = 0;
            cantidadPendienteNesting = 0;

            lblNestingSeleccionado.Text = "Nesting: --";
            lblParteSeleccionada.Text = "Parte: --";
            lblRecursoSeleccionado.Text = "Recurso: --";
            lblNstRefValor.Text = "--";
            lblCantidadInfo.Text = "Nesting - Programado: 0 | Reportado: 0 | Pendiente: 0";
            txtCantidad.Text = "0";
            btnReportar.Enabled = false;
            btnActualizar.Enabled = false;
            lblCantidadTitulo.Text = "CANTIDAD A REPORTAR (Quantity de la parte)";

            // Limpiar Job label
            Control[] controls = panelPrincipal.Controls.Find("lblJobValor", true);
            if (controls.Length > 0)
            {
                ((Label)controls[0]).Text = "--";
            }

            panelPartes.Controls.Clear();
        }

        private void MostrarCargando(bool mostrar, string mensaje = "")
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool, string>(MostrarCargando), mostrar, mensaje);
                return;
            }

            lblCargando.Visible = mostrar;
            if (!string.IsNullOrEmpty(mensaje))
            {
                lblCargando.Text = mensaje;
            }
            Application.DoEvents();
        }

        // ============== MÉTODOS DE ESTILO ==============

        private void MostrarNotificacion(string mensaje, Color color)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string, Color>(MostrarNotificacion), mensaje, color);
                return;
            }

            Label lblNotificacion = new Label
            {
                Text = mensaje,
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = color,
                AutoSize = true,
                Padding = new Padding(10, 5, 10, 5),
                Location = new Point(this.Width - 350, 80),
                BorderStyle = BorderStyle.FixedSingle
            };

            this.Controls.Add(lblNotificacion);
            lblNotificacion.BringToFront();

            Timer timerNotif = new Timer();
            timerNotif.Interval = 3000;
            timerNotif.Tick += (s, e) =>
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => {
                        this.Controls.Remove(lblNotificacion);
                        lblNotificacion.Dispose();
                    }));
                }
                else
                {
                    this.Controls.Remove(lblNotificacion);
                    lblNotificacion.Dispose();
                }
                timerNotif.Stop();
                timerNotif.Dispose();
            };
            timerNotif.Start();
        }
    }
}