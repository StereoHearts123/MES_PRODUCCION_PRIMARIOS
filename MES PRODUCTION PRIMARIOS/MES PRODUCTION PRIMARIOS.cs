using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace LaserCuttingApp
{
    public partial class FormLaserCutting : Form
    {
        // ============== CONEXIONES A BASES DE DATOS ==============
        private readonly string connectionStringMartinRea = @"Data Source=10.241.1.27;Initial Catalog=MartinRea MJ;User Id=sa;Password=Mrea$ql123;Connection Timeout=30;Max Pool Size=120;";
        private readonly string connectionStringORD_PROD = @"Data Source=10.241.1.27;Initial Catalog=ORD_PROD;User Id=sa;Password=Mrea$ql123;Connection Timeout=30;Max Pool Size=120;";
        private readonly string connectionStringMES = @"Data Source=10.241.1.27;Initial Catalog=MES;User Id=sa;Password=Mrea$ql123;Connection Timeout=60;Max Pool Size=120;";
        private readonly string connectionStringMES_PRODUCTION = @"Data Source=10.241.1.27;Initial Catalog=MES_PRODUCTION;User Id=sa;Password=Mrea$ql123;Connection Timeout=60;Max Pool Size=120;";

        // ============== CONSTANTES ==============
        private const int NOTIFICACION_DURACION_MS = 3000;

        // ============== VARIABLES PRINCIPALES ==============
        private string cncSeleccionado = "";
        private string nestingSeleccionado = "";
        private string parteSeleccionada = "";
        private string recursoSeleccionado = "";
        private string catIdSeleccionado = "0";
        private string nstRefActual = "";
        private string mnORefActual = "";
        private int cantidadProduccionActual = 0;
        private int cantidadProgramadaNesting = 0;
        private int cantidadReportadaNesting = 0;
        private int cantidadPendienteNesting = 0;
        private bool isLoading = false;
        private string cnnIngresado = "";

        // Cache para recursos (incluye CAT_ID)
        private readonly ConcurrentDictionary<string, RecursoInfo> cacheRecursos = new ConcurrentDictionary<string, RecursoInfo>();

        // Cache para SUBRESOURCE_ID
        private readonly ConcurrentDictionary<string, int> cacheSubresourceId = new ConcurrentDictionary<string, int>();

        // Semaphore para control de concurrencia
        private readonly System.Threading.SemaphoreSlim semaphore = new System.Threading.SemaphoreSlim(1, 1);

        // ============== ESTRUCTURAS DE DATOS ==============
        private List<ParteInfo> partesActuales = new List<ParteInfo>();

        public class RecursoInfo
        {
            public string Recurso { get; set; }
            public string Descripcion { get; set; }
            public string CAT_ID { get; set; }

            public RecursoInfo()
            {
                Recurso = "";
                Descripcion = "";
                CAT_ID = "0";
            }
        }

        public class ParteInfo
        {
            public string MnORef { get; set; }
            public string OprID { get; set; }
            public string PrdRefDst { get; set; }
            public string Recurso { get; set; }
            public int Cantidad { get; set; }
            public string NstRef { get; set; }
            public string NestingNombre { get; set; }

            public ParteInfo()
            {
                MnORef = "";
                OprID = "";
                PrdRefDst = "";
                Recurso = "";
                NstRef = "";
                NestingNombre = "";
            }
        }

        // ============== CONTROLES DE LA INTERFAZ ==============
        private Panel panelCabecera;
        private Panel panelPrincipal;
        private Panel panelInferior;
        private Label lblTitulo;
        private Label lblCNCSeleccionado;
        private Label lblNestingSeleccionado;
        private Label lblRecursoSeleccionado;
        private Label lblParteSeleccionada;
        private Label lblCatIdValor;
        private Label lblNstRefValor;
        private Label lblCantidadInfo;
        private Label lblEstadoConexion;
        private Label lblFechaActual;
        private Label lblHoraActual;
        private FlowLayoutPanel panelPartes;
        private Button btnReportar;
        private Button btnLimpiar;
        private Button btnBorrar;
        private Timer timerReloj;
        private Label lblCargando;
        private TextBox txtCantidad;
        private Label lblCantidadTitulo;
        private Label lblCNCIngresado;
        private TextBox txtCNCDisplay;
        private Button btnConfirmarCNC;

        // Botones del PIN PAD
        private Button btn0, btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn9;

        private Button parteSeleccionadaBtn = null;

        // ============== COLORES ==============
        private readonly Color colorPrimario = Color.FromArgb(106, 13, 18);
        private readonly Color colorSecundario = Color.FromArgb(0, 114, 198);
        private readonly Color colorExito = Color.FromArgb(40, 167, 69);
        private readonly Color colorPeligro = Color.FromArgb(220, 53, 69);
        private readonly Color colorInfo = Color.FromArgb(23, 162, 184);
        private readonly Color colorFondo = Color.FromArgb(245, 245, 245);
        private readonly Color colorSeleccionado = Color.FromArgb(0, 114, 198);
        private readonly Color colorBorde = Color.FromArgb(200, 200, 200);

        public FormLaserCutting()
        {
            InitializeComponent();
            ConfigurarEstiloGeneral();
            ConfigurarInterfaz();
            _ = CargarDatosInicialesAsync();
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
            this.Text = "SISTEMA DE REPORTE - CORTADORAS LASER";
            this.BackColor = colorFondo;
            this.DoubleBuffered = true;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.Size = new Size(1366, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
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
                Text = "REPORTE DE PRODUCCION - CORTADORAS LASER",
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

            ConfigurarPanelCNC();
            ConfigurarPanelPartes();
            ConfigurarPanelControl();

            this.Controls.Add(panelPrincipal);
        }

        private void ConfigurarPanelCNC()
        {
            Panel panelCNC = new Panel
            {
                Location = new Point(15, 80),
                Size = new Size(430, 600),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            Label lblTituloCNC = new Label
            {
                Text = "INGRESE CODIGO CNC",
                Font = new Font("Arial", 14, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(15, 15),
                AutoSize = true
            };

            // Display del CNC
            txtCNCDisplay = new TextBox
            {
                Location = new Point(15, 50),
                Size = new Size(400, 50),
                Font = new Font("Arial", 18, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
                ReadOnly = true,
                Text = ""
            };

            lblCNCIngresado = new Label
            {
                Text = "CNC: --",
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = colorPrimario,
                Location = new Point(15, 110),
                AutoSize = true,
                Visible = false
            };

            // Panel para botones del PIN PAD - AUMENTAR EL TAMAÑO
            Panel panelBotones = new Panel
            {
                Location = new Point(15, 140),
                Size = new Size(450, 390), // Cambiado de 200 a 350
                BackColor = Color.White    // Agregar color de fondo para depuración
            };

            // Crear botones del PIN PAD
            int btnWidth = 120;
            int btnHeight = 80;
            int startX = 10;
            int startY = 10;
            int spacing = 15;

            // Fila 1: 1, 2, 3
            btn1 = CrearBotonPinPad("1", startX, startY, btnWidth, btnHeight);
            btn2 = CrearBotonPinPad("2", startX + btnWidth + spacing, startY, btnWidth, btnHeight);
            btn3 = CrearBotonPinPad("3", startX + (btnWidth + spacing) * 2, startY, btnWidth, btnHeight);

            // Fila 2: 4, 5, 6
            btn4 = CrearBotonPinPad("4", startX, startY + btnHeight + spacing, btnWidth, btnHeight);
            btn5 = CrearBotonPinPad("5", startX + btnWidth + spacing, startY + btnHeight + spacing, btnWidth, btnHeight);
            btn6 = CrearBotonPinPad("6", startX + (btnWidth + spacing) * 2, startY + btnHeight + spacing, btnWidth, btnHeight);

            // Fila 3: 7, 8, 9
            btn7 = CrearBotonPinPad("7", startX, startY + (btnHeight + spacing) * 2, btnWidth, btnHeight);
            btn8 = CrearBotonPinPad("8", startX + btnWidth + spacing, startY + (btnHeight + spacing) * 2, btnWidth, btnHeight);
            btn9 = CrearBotonPinPad("9", startX + (btnWidth + spacing) * 2, startY + (btnHeight + spacing) * 2, btnWidth, btnHeight);

            // Fila 4: 0, Borrar, Confirmar
            btn0 = CrearBotonPinPad("0", startX + btnWidth + spacing, startY + (btnHeight + spacing) * 3, btnWidth, btnHeight);
            btnBorrar = new Button
            {
                Text = "⌫",
                Location = new Point(startX, startY + (btnHeight + spacing) * 3),
                Size = new Size(btnWidth, btnHeight),
                Font = new Font("Arial", 20, FontStyle.Bold),
                BackColor = Color.Orange,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnBorrar.FlatAppearance.BorderSize = 0;
            btnBorrar.Click += BtnBorrar_Click;

            btnConfirmarCNC = new Button
            {
                Text = "✓",
                Location = new Point(startX + (btnWidth + spacing) * 2, startY + (btnHeight + spacing) * 3),
                Size = new Size(btnWidth, btnHeight),
                Font = new Font("Arial", 20, FontStyle.Bold),
                BackColor = colorExito,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnConfirmarCNC.FlatAppearance.BorderSize = 0;
            btnConfirmarCNC.Click += async (s, e) => await ConfirmarCNCAsync();

            panelBotones.Controls.AddRange(new Control[] { btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn9, btn0, btnBorrar, btnConfirmarCNC });
            panelCNC.Controls.AddRange(new Control[] { lblTituloCNC, txtCNCDisplay, lblCNCIngresado, panelBotones });
            panelPrincipal.Controls.Add(panelCNC);
        }

        private Button CrearBotonPinPad(string texto, int x, int y, int width, int height)
        {
            Button btn = new Button
            {
                Text = texto,
                Location = new Point(x, y),
                Size = new Size(width, height),
                Font = new Font("Arial", 18, FontStyle.Bold),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.BorderColor = colorBorde;
            btn.Click += BtnPinPad_Click;
            return btn;
        }

        private void BtnPinPad_Click(object sender, EventArgs e)
        {
            if (sender is Button btn)
            {
                cnnIngresado += btn.Text;
                txtCNCDisplay.Text = cnnIngresado;
            }
        }

        private void BtnBorrar_Click(object sender, EventArgs e)
        {
            if (cnnIngresado.Length > 0)
            {
                cnnIngresado = cnnIngresado.Substring(0, cnnIngresado.Length - 1);
                txtCNCDisplay.Text = cnnIngresado;
            }
        }

        private void ConfigurarPanelPartes()
        {
            Panel panelPartesContainer = new Panel
            {
                Location = new Point(460, 80),
                Size = new Size(430, 600),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            Label lblPartesTitulo = new Label
            {
                Text = "NUMEROS DE PARTE",
                Font = new Font("Arial", 14, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(15, 15),
                AutoSize = true
            };

            panelPartes = new FlowLayoutPanel
            {
                Location = new Point(10, 60),
                Size = new Size(405, 600),
                AutoScroll = true,
                Padding = new Padding(5),
                WrapContents = true,
                BackColor = Color.White
            };

            panelPartesContainer.Controls.AddRange(new Control[] { lblPartesTitulo, panelPartes });
            panelPrincipal.Controls.Add(panelPartesContainer);
        }

        private void ConfigurarPanelControl()
        {
            Panel panelControl = new Panel
            {
                Location = new Point(905, 80),
                Size = new Size(440, 600),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            Panel panelInfoTrabajo = new Panel
            {
                Location = new Point(15,20),
                Size = new Size(410, 210),
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };

            lblCNCSeleccionado = new Label
            {
                Text = "CNC: --",
                Font = new Font("Arial", 12, FontStyle.Bold),
                Location = new Point(10, 10),
                AutoSize = true
            };

            lblNestingSeleccionado = new Label
            {
                Text = "Trabajo: --",
                Font = new Font("Arial", 11, FontStyle.Bold),
                Location = new Point(10, 40),
                AutoSize = true
            };

            lblRecursoSeleccionado = new Label
            {
                Text = "Recurso: --",
                Font = new Font("Arial", 11),
                Location = new Point(10, 70),
                AutoSize = true
            };

            lblParteSeleccionada = new Label
            {
                Text = "Parte: --",
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(10, 100),
                AutoSize = true
            };

            Label lblCatIdTitulo = new Label
            {
                Text = "CAT_ID:",
                Font = new Font("Arial", 11, FontStyle.Bold),
                Location = new Point(10, 130),
                AutoSize = true
            };

            lblCatIdValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorExito,
                Location = new Point(80, 130),
                AutoSize = true
            };

            Label lblJobInfo = new Label
            {
                Text = "JOB / Operacion:",
                Font = new Font("Arial", 11, FontStyle.Bold),
                Location = new Point(10, 160),
                AutoSize = true
            };

            Label lblJobValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 10),
                ForeColor = colorInfo,
                Location = new Point(140, 160),
                AutoSize = true,
                Name = "lblJobValor"
            };

            lblNstRefValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 9),
                ForeColor = Color.Gray,
                Location = new Point(10, 185),
                AutoSize = true
            };

            panelInfoTrabajo.Controls.AddRange(new Control[] { lblCNCSeleccionado, lblNestingSeleccionado, lblRecursoSeleccionado,
                lblParteSeleccionada, lblCatIdTitulo, lblCatIdValor, lblJobInfo, lblJobValor, lblNstRefValor });

            lblCantidadInfo = new Label
            {
                Text = "Programado: 0 | Reportado: 0 | Pendiente: 0",
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = colorPrimario,
                Location = new Point(15, 270),
                AutoSize = true
            };

            lblCantidadTitulo = new Label
            {
                Text = "CANTIDAD A REPORTAR",
                Font = new Font("Arial", 12, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(15, 310),
                AutoSize = true
            };

            txtCantidad = new TextBox
            {
                Location = new Point(15, 340),
                Size = new Size(410, 50),
                Font = new Font("Arial", 20, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
                Text = "0",
                ReadOnly = true
            };

            btnReportar = new Button
            {
                Text = "REPORTAR PRODUCCION",
                Location = new Point(15, 420),
                Size = new Size(410, 55),
                Font = new Font("Arial", 12, FontStyle.Bold),
                BackColor = colorExito,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnReportar.FlatAppearance.BorderSize = 0;
            btnReportar.Click += async (s, e) => await ReportarProduccionAsync();

            btnLimpiar = new Button
            {
                Text = "LIMPIAR SELECCION",
                Location = new Point(15, 490),
                Size = new Size(410, 45),
                Font = new Font("Arial", 11, FontStyle.Bold),
                BackColor = colorPeligro,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnLimpiar.FlatAppearance.BorderSize = 0;
            btnLimpiar.Click += (s, e) => LimpiarSeleccion();

            lblCargando = new Label
            {
                Text = "Cargando datos...",
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorInfo,
                Location = new Point(15, 560),
                AutoSize = true,
                Visible = false
            };

            panelControl.Controls.AddRange(new Control[] { panelInfoTrabajo, lblCantidadInfo, lblCantidadTitulo,
                txtCantidad, btnReportar, btnLimpiar, lblCargando });
            panelPrincipal.Controls.Add(panelControl);
        }

        private void ConfigurarPanelInferior()
        {
            panelInferior = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,    
                BackColor = Color.FromArgb(240, 240, 240),
                BorderStyle = BorderStyle.FixedSingle
            };

            lblEstadoConexion = new Label
            {
                Text = "CONECTADO",
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
                Location = new Point(200, 28),
                AutoSize = true
            };

            panelInferior.Controls.AddRange(new Control[] { lblEstadoConexion, lblFechaActual, lblHoraActual });
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

        private async Task CargarDatosInicialesAsync()
        {
            try
            {
                await ProbarConexionAsync();
                MostrarNotificacion("Sistema listo. Ingrese un codigo CNC.", Color.Green);
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al conectar: {ex.Message}", Color.Red);
                lblEstadoConexion.Text = "DESCONECTADO";
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

        // ============== OBTENER RECURSO Y CAT_ID DESDE capability ==============
        private async Task<RecursoInfo> ObtenerRecursoDesdeTablaAsync(string numeroParte)
        {
            if (cacheRecursos.TryGetValue(numeroParte, out RecursoInfo cachedInfo))
                return cachedInfo;

            try
            {
                string query = @"
                SELECT TOP 1 
                    ISNULL(c.AORESC, 'No especificado') as recurso,
                    ISNULL(l.desc_recurso, '') as desc_recurso,
                    ISNULL(c.AVCATA, '0') as cat_id
                FROM [MES_PRODUCTION].[dbo].[capability] c
                LEFT JOIN [ORD_PROD].[dbo].[laser_recursos] l 
                    ON c.AOPART COLLATE Modern_Spanish_CI_AS = l.part COLLATE Modern_Spanish_CI_AS
                WHERE c.AOPART = @numeroParte
                AND c.ARREPP = 'Y'";

                DataTable dt = new DataTable();

                await Task.Run(() =>
                {
                    using (SqlConnection conn = new SqlConnection(connectionStringMES_PRODUCTION))
                    {
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@numeroParte", numeroParte);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                });

                RecursoInfo recursoInfo = new RecursoInfo();

                if (dt.Rows.Count > 0)
                {
                    recursoInfo.Recurso = dt.Rows[0]["recurso"].ToString() ?? "No especificado";
                    recursoInfo.Descripcion = dt.Rows[0]["desc_recurso"].ToString() ?? "";
                    recursoInfo.CAT_ID = dt.Rows[0]["cat_id"].ToString() ?? "0";
                }
                else
                {
                    recursoInfo.Recurso = "No especificado";
                    recursoInfo.Descripcion = "";
                    recursoInfo.CAT_ID = "0";
                }

                cacheRecursos[numeroParte] = recursoInfo;
                return recursoInfo;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al obtener recurso: {ex.Message}");
                return new RecursoInfo { Recurso = "Error", Descripcion = "", CAT_ID = "0" };
            }
        }

        // ============== OBTENER SUBRESOURCE_ID DESDE MES.TBL_SUBRESOURCE ==============
        private async Task<int> ObtenerSubresourceIdAsync(string recurso)
        {
            if (string.IsNullOrEmpty(recurso) || recurso == "No especificado" || recurso == "Error")
                return 0;

            if (cacheSubresourceId.TryGetValue(recurso, out int cachedId))
                return cachedId;

            try
            {
                string query = @"
                SELECT TOP 1 ID 
                FROM MES.TBL_SUBRESOURCE 
                WHERE SUB_NAME LIKE @recurso
                AND ACTIVE = 1
                ORDER BY ID";

                using (SqlConnection conn = new SqlConnection(connectionStringMES))
                {
                    await conn.OpenAsync();
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@recurso", "%" + recurso + "%");
                        object result = await cmd.ExecuteScalarAsync();

                        if (result != null && result != DBNull.Value)
                        {
                            int id = Convert.ToInt32(result);
                            cacheSubresourceId[recurso] = id;
                            return id;
                        }
                    }
                }

                string queryDefault = "SELECT TOP 1 ID FROM MES.TBL_SUBRESOURCE WHERE ACTIVE = 1 ORDER BY ID";
                using (SqlConnection conn = new SqlConnection(connectionStringMES))
                {
                    await conn.OpenAsync();
                    using (SqlCommand cmd = new SqlCommand(queryDefault, conn))
                    {
                        object result = await cmd.ExecuteScalarAsync();
                        if (result != null && result != DBNull.Value)
                        {
                            int id = Convert.ToInt32(result);
                            cacheSubresourceId[recurso] = id;
                            return id;
                        }
                    }
                }

                return 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error obteniendo SUBRESOURCE_ID: {ex.Message}");
                return 0;
            }
        }

        // ============== INSERTAR EN MES.TBL_PRODUCTION ==============
        private async Task InsertarEnTBLProductionAsync(int subresourceId, int catId, int cantidad)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionStringMES))
                {
                    await conn.OpenAsync();

                    string query = @"
                    INSERT INTO MES.TBL_PRODUCTION 
                    (SUBRESOURCE_ID, CATID, QTY, TIMESTAMP)
                    VALUES (@subresource_id, @catid, @qty, @timestamp)";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@subresource_id", subresourceId > 0 ? (object)subresourceId : DBNull.Value);
                        cmd.Parameters.AddWithValue("@catid", catId);
                        cmd.Parameters.AddWithValue("@qty", cantidad);
                        cmd.Parameters.AddWithValue("@timestamp", DateTime.Now);

                        await cmd.ExecuteNonQueryAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en MES.TBL_PRODUCTION: {ex.Message}");
                throw;
            }
        }

        // ============== GUARDAR EN TABLA EXISTENTE TBL_MES_MARS_lASER ==============
        private async Task GuardarEnTablaLaserAsync(int cantidad)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionStringMES_PRODUCTION))
                {
                    await conn.OpenAsync();

                    string createTable = @"
                    IF NOT EXISTS (SELECT * FROM sys.objects WHERE name = 'TBL_MES_MARS_lASER' AND type = 'U')
                    CREATE TABLE TBL_MES_MARS_lASER (
                        ID INT IDENTITY(1,1) PRIMARY KEY,
                        CAT_ID NVARCHAR(50) NULL,
                        RESOURCE NVARCHAR(200) NULL,
                        DATATIME DATETIME NULL,
                        [COUNT] INT NULL
                    )";

                    using (SqlCommand cmd = new SqlCommand(createTable, conn))
                    {
                        await cmd.ExecuteNonQueryAsync();
                    }

                    string insertSql = @"
                    INSERT INTO TBL_MES_MARS_lASER (CAT_ID, RESOURCE, DATATIME, [COUNT])
                    VALUES (@cat_id, @resource, @datatime, @count)";

                    using (SqlCommand cmd = new SqlCommand(insertSql, conn))
                    {
                        cmd.Parameters.AddWithValue("@cat_id", catIdSeleccionado);
                        cmd.Parameters.AddWithValue("@resource", recursoSeleccionado);
                        cmd.Parameters.AddWithValue("@datatime", DateTime.Now);
                        cmd.Parameters.AddWithValue("@count", cantidad);

                        await cmd.ExecuteNonQueryAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en TBL_MES_MARS_lASER: {ex.Message}");
                throw;
            }
        }

        private async Task ConfirmarCNCAsync()
        {
            if (string.IsNullOrWhiteSpace(cnnIngresado))
            {
                MostrarNotificacion("Ingrese un codigo CNC", Color.Orange);
                return;
            }

            if (isLoading) return;

            await semaphore.WaitAsync();
            try
            {
                isLoading = true;
                MostrarCargando(true, $"Buscando CNC: {cnnIngresado}...");

                // Buscar todas las partes de todos los nestings del CNC
                string queryPartes = @"
                SELECT 
                    n.NstRef,
                    n.Name as NestingNombre,
                    p.MnORef,
                    p.OprID,
                    p.PrdRefDst,
                    p.Quantity,
                    n.Quantity as Cantidad_Programada,
                    ISNULL(n.CutQuantity, 0) as Cantidad_Reportada,
                    ISNULL(n.Quantity, 0) - ISNULL(n.CutQuantity, 0) as Cantidad_Pendiente
                FROM [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000100] n
                INNER JOIN [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000500] p 
                    ON n.NstRef = p.NstRef
                WHERE n.CNC = @cnc 
                    AND p.PrdRefDst IS NOT NULL 
                    AND p.PrdRefDst != ''
                ORDER BY n.Name, p.PrdRefDst";

                DataTable dt = new DataTable();

                await Task.Run(() =>
                {
                    using (SqlConnection conn = new SqlConnection(connectionStringMartinRea))
                    {
                        using (SqlCommand cmd = new SqlCommand(queryPartes, conn))
                        {
                            cmd.Parameters.AddWithValue("@cnc", cnnIngresado);
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt);
                            }
                        }
                    }
                });

                partesActuales.Clear();

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        partesActuales.Add(new ParteInfo
                        {
                            NstRef = row["NstRef"].ToString() ?? "",
                            NestingNombre = row["NestingNombre"].ToString() ?? "",
                            MnORef = row["MnORef"].ToString() ?? "",
                            OprID = row["OprID"].ToString() ?? "",
                            PrdRefDst = row["PrdRefDst"].ToString() ?? "",
                            Cantidad = Convert.ToInt32(row["Quantity"]),
                            Recurso = row["NestingNombre"].ToString() ?? ""
                        });
                    }

                    cncSeleccionado = cnnIngresado;
                    lblCNCSeleccionado.Text = $"CNC: {cncSeleccionado}";
                    lblCNCIngresado.Text = $"CNC: {cncSeleccionado}";
                    lblCNCIngresado.Visible = true;

                    // Crear botones de partes
                    CrearBotonesPartes();

                    // Seleccionar automáticamente la primera parte
                    if (partesActuales.Count > 0)
                    {
                        var primeraParte = partesActuales.First();

                        foreach (Control control in panelPartes.Controls)
                        {
                            if (control is Button btn && btn.Tag is ParteInfo parte && parte.PrdRefDst == primeraParte.PrdRefDst)
                            {
                                btn.BackColor = colorSeleccionado;
                                btn.ForeColor = Color.White;
                                parteSeleccionadaBtn = btn;
                                await ActualizarInfoParteAsync(primeraParte);
                                break;
                            }
                        }
                    }

                    MostrarNotificacion($"CNC {cnnIngresado} encontrado. {partesActuales.Count} parte(s) cargada(s).", Color.Green);
                }
                else
                {
                    MostrarNotificacion($"No se encontraron partes para el CNC: {cnnIngresado}", Color.Orange);
                    LimpiarSeleccion();
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
                semaphore.Release();
            }
        }

        private void CrearBotonesPartes()
        {
            panelPartes.SuspendLayout();
            panelPartes.Controls.Clear();

            int index = 1;
            foreach (var parte in partesActuales)
            {
                Button btnParte = new Button
                {
                    Text = $"{index:00}. {parte.PrdRefDst}\nCantidad: {parte.Cantidad:N0}\n{parte.NestingNombre}",
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    BackColor = Color.White,
                    Size = new Size(400, 65),
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand,
                    Margin = new Padding(3),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Tag = parte,
                    UseVisualStyleBackColor = true
                };
                btnParte.FlatAppearance.BorderSize = 1;
                btnParte.FlatAppearance.BorderColor = colorBorde;
                btnParte.Click += BtnParte_Click;
                panelPartes.Controls.Add(btnParte);
                index++;
            }
            panelPartes.ResumeLayout();
        }

        private async void BtnParte_Click(object sender, EventArgs e)
        {
            if (sender is Button btn && btn.Tag is ParteInfo parte)
            {
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
        }

        private async Task ActualizarInfoParteAsync(ParteInfo parte)
        {
            try
            {
                RecursoInfo recursoInfo = await ObtenerRecursoDesdeTablaAsync(parte.PrdRefDst);

                catIdSeleccionado = recursoInfo.CAT_ID;

                if (lblCatIdValor != null)
                {
                    lblCatIdValor.Text = catIdSeleccionado;
                    lblCatIdValor.ForeColor = catIdSeleccionado != "0" ? Color.Green : Color.Gray;
                }

                string queryNesting = @"
                SELECT Name as Recurso, Quantity as Cantidad_Programada, 
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

                    if (recursoInfo.Recurso != "No especificado" && recursoInfo.Recurso != "Error")
                    {
                        recursoSeleccionado = recursoInfo.Recurso;
                        lblRecursoSeleccionado.Text = string.IsNullOrEmpty(recursoInfo.Descripcion)
                            ? $"Recurso: {recursoSeleccionado}"
                            : $"Recurso: {recursoSeleccionado} - {recursoInfo.Descripcion}";
                    }
                    else
                    {
                        recursoSeleccionado = row["Recurso"].ToString() ?? "";
                        lblRecursoSeleccionado.Text = $"Recurso: {recursoSeleccionado}";
                    }

                    cantidadProgramadaNesting = Convert.ToInt32(row["Cantidad_Programada"]);
                    cantidadReportadaNesting = Convert.ToInt32(row["Cantidad_Reportada"]);
                    cantidadPendienteNesting = Convert.ToInt32(row["Cantidad_Pendiente"]);
                    nstRefActual = parte.NstRef;
                    mnORefActual = parte.MnORef;
                    parteSeleccionada = parte.PrdRefDst;
                    cantidadProduccionActual = parte.Cantidad;
                    nestingSeleccionado = parte.NestingNombre;

                    lblParteSeleccionada.Text = $"Parte: {parteSeleccionada}";
                    lblNestingSeleccionado.Text = $"Trabajo: {nestingSeleccionado}";

                    Control[] controls = panelPrincipal.Controls.Find("lblJobValor", true);
                    if (controls.Length > 0 && controls[0] is Label lblJob)
                        lblJob.Text = $"{parte.MnORef} / {parte.OprID}";

                    lblNstRefValor.Text = $"NstRef: {nstRefActual}";
                    lblCantidadInfo.Text = $"Programado: {cantidadProgramadaNesting:N0} | Reportado: {cantidadReportadaNesting:N0} | Pendiente: {cantidadPendienteNesting:N0}";
                    txtCantidad.Text = cantidadProduccionActual.ToString("N0");
                    txtCantidad.ForeColor = cantidadProduccionActual > 0 ? Color.Green : Color.Gray;

                    btnReportar.Enabled = cantidadProduccionActual > 0 && catIdSeleccionado != "0";

                    if (catIdSeleccionado == "0")
                    {
                        MostrarNotificacion($"La parte '{parteSeleccionada}' no tiene CAT_ID asignado", Color.Orange);
                    }
                }
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al obtener informacion: {ex.Message}", Color.Red);
            }
        }

        private int ObtenerCantidadIngresada()
        {
            if (string.IsNullOrEmpty(txtCantidad.Text)) return 0;
            return int.TryParse(txtCantidad.Text.Replace(",", ""), out int resultado) ? resultado : 0;
        }

        private async Task ReportarProduccionAsync()
        {
            if (string.IsNullOrEmpty(nstRefActual))
            {
                MostrarNotificacion("No se encontro informacion para esta parte", Color.Red);
                return;
            }

            if (catIdSeleccionado == "0" || string.IsNullOrEmpty(catIdSeleccionado))
            {
                MostrarNotificacion("La parte seleccionada no tiene un CAT_ID valido", Color.Red);
                return;
            }

            int cantidadReportar = ObtenerCantidadIngresada();

            if (cantidadReportar <= 0)
            {
                MostrarNotificacion("La cantidad debe ser mayor a 0", Color.Red);
                return;
            }

            DialogResult result = MessageBox.Show(
                $"Confirmar reporte de {cantidadReportar:N0} piezas?\n\n" +
                $"CNC: {cncSeleccionado}\n" +
                $"Trabajo: {nestingSeleccionado}\n" +
                $"Recurso: {recursoSeleccionado}\n" +
                $"Parte: {parteSeleccionada}\n" +
                $"CAT_ID: {catIdSeleccionado}\n" +
                $"JOB: {mnORefActual}",
                "Confirmar Reporte",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                await semaphore.WaitAsync();
                try
                {
                    btnReportar.Enabled = false;
                    btnReportar.Text = "PROCESANDO...";

                    await GuardarEnTablaLaserAsync(cantidadReportar);
                    int subresourceId = await ObtenerSubresourceIdAsync(recursoSeleccionado);
                    int catIdInt = int.Parse(catIdSeleccionado);
                    await InsertarEnTBLProductionAsync(subresourceId, catIdInt, cantidadReportar);

                    MostrarNotificacion($"Reporte exitoso: {cantidadReportar:N0} piezas", Color.Green);
                }
                catch (Exception ex)
                {
                    MostrarNotificacion($"Error al reportar: {ex.Message}", Color.Red);
                }
                finally
                {
                    btnReportar.Text = "REPORTAR PRODUCCION";
                    btnReportar.Enabled = true;
                    semaphore.Release();
                }
            }
        }

        private void LimpiarSeleccion()
        {
            if (parteSeleccionadaBtn != null)
            {
                parteSeleccionadaBtn.BackColor = Color.White;
                parteSeleccionadaBtn.ForeColor = Color.Black;
                parteSeleccionadaBtn = null;
            }

            cnnIngresado = "";
            txtCNCDisplay.Text = "";
            cncSeleccionado = "";
            nestingSeleccionado = "";
            parteSeleccionada = "";
            recursoSeleccionado = "";
            catIdSeleccionado = "0";
            nstRefActual = "";
            mnORefActual = "";
            cantidadProduccionActual = 0;
            partesActuales.Clear();

            lblCNCSeleccionado.Text = "CNC: --";
            lblNestingSeleccionado.Text = "Trabajo: --";
            lblParteSeleccionada.Text = "Parte: --";
            lblRecursoSeleccionado.Text = "Recurso: --";
            if (lblCatIdValor != null) lblCatIdValor.Text = "--";
            lblNstRefValor.Text = "--";
            lblCantidadInfo.Text = "Programado: 0 | Reportado: 0 | Pendiente: 0";
            txtCantidad.Text = "0";
            btnReportar.Enabled = false;
            lblCNCIngresado.Visible = false;

            Control[] controls = panelPrincipal.Controls.Find("lblJobValor", true);
            if (controls.Length > 0 && controls[0] is Label lblJob)
                lblJob.Text = "--";

            panelPartes.Controls.Clear();
        }

        private void MostrarCargando(bool mostrar, string mensaje = "")
        {
            if (InvokeRequired)
            {
                Invoke(new Action<bool, string>(MostrarCargando), mostrar, mensaje);
                return;
            }

            lblCargando.Visible = mostrar;
            if (!string.IsNullOrEmpty(mensaje))
                lblCargando.Text = mensaje;
            Application.DoEvents();
        }

        private void MostrarNotificacion(string mensaje, Color color)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, Color>(MostrarNotificacion), mensaje, color);
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
                Location = new Point(this.Width - 650, 80),
                BorderStyle = BorderStyle.FixedSingle
            };

            this.Controls.Add(lblNotificacion);
            lblNotificacion.BringToFront();

            Timer timerNotif = new Timer();
            timerNotif.Interval = NOTIFICACION_DURACION_MS;
            timerNotif.Tick += (s, e) =>
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(() => {
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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            semaphore?.Dispose();
            timerReloj?.Dispose();
            base.OnFormClosing(e);
        }
    }
}