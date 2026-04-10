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

        // Cache para recursos (incluye CAT_ID)
        private readonly ConcurrentDictionary<string, RecursoInfo> cacheRecursos = new ConcurrentDictionary<string, RecursoInfo>();

        // Cache para SUBRESOURCE_ID
        private readonly ConcurrentDictionary<string, int> cacheSubresourceId = new ConcurrentDictionary<string, int>();

        // Semaphore para control de concurrencia
        private readonly System.Threading.SemaphoreSlim semaphore = new System.Threading.SemaphoreSlim(1, 1);

        // ============== ESTRUCTURAS DE DATOS ==============
        private List<NestingInfo> nestingsActuales = new List<NestingInfo>();
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

        public class NestingInfo
        {
            public string NstRef { get; set; }
            public string Nombre { get; set; }
            public int CantidadProgramada { get; set; }
            public int CantidadReportada { get; set; }
            public int CantidadPendiente { get; set; }
            public string Estado { get; set; }

            public NestingInfo()
            {
                NstRef = "";
                Nombre = "";
                Estado = "";
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

            public ParteInfo()
            {
                MnORef = "";
                OprID = "";
                PrdRefDst = "";
                Recurso = "";
                NstRef = "";
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
        private TextBox txtBuscarCNC;
        private FlowLayoutPanel panelNestings;
        private FlowLayoutPanel panelPartes;
        private Button btnBuscar;
        private Button btnReportar;
        private Button btnLimpiar;
        private Timer timerReloj;
        private Label lblCargando;
        private TextBox txtCantidad;
        private Label lblCantidadTitulo;

        private Button nestingSeleccionadoBtn = null;
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

            ConfigurarPanelBusqueda();
            ConfigurarPanelNestings();
            ConfigurarPanelPartes();
            ConfigurarPanelControl();

            this.Controls.Add(panelPrincipal);
        }

        private void ConfigurarPanelBusqueda()
        {
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
                Text = "Ingrese codigo CNC...",
                ForeColor = Color.Gray
            };
            txtBuscarCNC.GotFocus += TxtBuscarCNC_GotFocus;
            txtBuscarCNC.LostFocus += TxtBuscarCNC_LostFocus;
            txtBuscarCNC.KeyPress += TxtBuscarCNC_KeyPress;

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

            lblCargando = new Label
            {
                Text = "Cargando datos...",
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorInfo,
                Location = new Point(15, 125),
                AutoSize = true,
                Visible = false
            };

            panelBusqueda.Controls.AddRange(new Control[] { lblBuscarCNC, txtBuscarCNC, btnBuscar, btnRefrescar, lblEjemplo });
            panelPrincipal.Controls.AddRange(new Control[] { panelBusqueda, lblCargando });
        }

        private void ConfigurarPanelNestings()
        {
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

            panelNestings = new FlowLayoutPanel
            {
                Location = new Point(10, 60),
                Size = new Size(405, 445),
                AutoScroll = true,
                Padding = new Padding(5),
                WrapContents = true,
                BackColor = Color.White
            };

            panelNestingsContainer.Controls.AddRange(new Control[] { lblNestingsTitulo, panelNestings });
            panelPrincipal.Controls.Add(panelNestingsContainer);
        }

        private void ConfigurarPanelPartes()
        {
            Panel panelPartesContainer = new Panel
            {
                Location = new Point(460, 160),
                Size = new Size(430, 520),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            Label lblPartesTitulo = new Label
            {
                Text = "NUMEROS DE PARTE",
                Font = new Font("Arial", 12, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(10, 10),
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

            panelPartesContainer.Controls.AddRange(new Control[] { lblPartesTitulo, panelPartes });
            panelPrincipal.Controls.Add(panelPartesContainer);
        }

        private void ConfigurarPanelControl()
        {
            Panel panelControl = new Panel
            {
                Location = new Point(905, 160),
                Size = new Size(440, 520),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };

            Panel panelInfoTrabajo = new Panel
            {
                Location = new Point(15, 40),
                Size = new Size(410, 180),
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

            Label lblCatIdTitulo = new Label
            {
                Text = "CAT_ID:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 110),
                AutoSize = true
            };

            lblCatIdValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = colorExito,
                Location = new Point(70, 110),
                AutoSize = true
            };

            Label lblJobInfo = new Label
            {
                Text = "JOB / Operacion:",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Location = new Point(10, 135),
                AutoSize = true
            };

            Label lblJobValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = colorInfo,
                Location = new Point(130, 135),
                AutoSize = true,
                Name = "lblJobValor"
            };

            lblNstRefValor = new Label
            {
                Text = "--",
                Font = new Font("Arial", 10, FontStyle.Bold),
                ForeColor = colorInfo,
                Location = new Point(15, 160),
                AutoSize = true
            };

            panelInfoTrabajo.Controls.AddRange(new Control[] { lblCNCSeleccionado, lblNestingSeleccionado, lblRecursoSeleccionado,
                lblParteSeleccionada, lblCatIdTitulo, lblCatIdValor, lblJobInfo, lblJobValor, lblNstRefValor });

            lblCantidadInfo = new Label
            {
                Text = "Programado: 0 | Reportado: 0 | Pendiente: 0",
                Font = new Font("Arial", 9),
                ForeColor = colorPrimario,
                Location = new Point(15, 240),
                AutoSize = true
            };

            lblCantidadTitulo = new Label
            {
                Text = "CANTIDAD A REPORTAR",
                Font = new Font("Arial", 11, FontStyle.Bold),
                ForeColor = colorSecundario,
                Location = new Point(15, 270),
                AutoSize = true
            };

            txtCantidad = new TextBox
            {
                Location = new Point(15, 300),
                Size = new Size(410, 45),
                Font = new Font("Arial", 18, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
                Text = "0",
                ReadOnly = true
            };

            btnReportar = new Button
            {
                Text = "MANDAR A REPORTAR PRODUCCION",
                Location = new Point(15, 360),
                Size = new Size(410, 50),
                Font = new Font("Arial", 11, FontStyle.Bold),
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
                Location = new Point(15, 420),
                Size = new Size(410, 45),
                Font = new Font("Arial", 11, FontStyle.Bold),
                BackColor = colorPeligro,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnLimpiar.FlatAppearance.BorderSize = 0;
            btnLimpiar.Click += (s, e) => LimpiarSeleccion();

            panelControl.Controls.AddRange(new Control[] { panelInfoTrabajo, lblCantidadInfo, lblCantidadTitulo,
                txtCantidad, btnReportar, btnLimpiar });
            panelPrincipal.Controls.Add(panelControl);
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

        private void TxtBuscarCNC_GotFocus(object sender, EventArgs e)
        {
            if (txtBuscarCNC.Text == "Ingrese codigo CNC...")
            {
                txtBuscarCNC.Text = "";
                txtBuscarCNC.ForeColor = Color.Black;
            }
        }

        private void TxtBuscarCNC_LostFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBuscarCNC.Text))
            {
                txtBuscarCNC.Text = "Ingrese codigo CNC...";
                txtBuscarCNC.ForeColor = Color.Gray;
            }
        }

        private void TxtBuscarCNC_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                _ = BuscarCNCAsync();
            }
        }

        private async Task CargarDatosInicialesAsync()
        {
            try
            {
                await ProbarConexionAsync();
                MostrarNotificacion("Sistema listo. Ingrese un codigo CNC para buscar.", Color.Green);
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

                // Si no encuentra, intentar obtener el primer activo
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

        // ============== INSERTAR EN MES.TBL_PRODUCTION (historico hora por hora) ==============
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

                        int rows = await cmd.ExecuteNonQueryAsync();
                        System.Diagnostics.Debug.WriteLine($"Insertado en MES.TBL_PRODUCTION: SubresourceId={subresourceId}, CATID={catId}, QTY={cantidad}, Filas={rows}");
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
                        System.Diagnostics.Debug.WriteLine($"Insertado en TBL_MES_MARS_lASER: CAT_ID={catIdSeleccionado}, Recurso={recursoSeleccionado}, Cantidad={cantidad}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en TBL_MES_MARS_lASER: {ex.Message}");
                throw;
            }
        }

        private async Task BuscarCNCAsync()
        {
            string cncABuscar = txtBuscarCNC.Text;
            if (cncABuscar == "Ingrese codigo CNC..." || string.IsNullOrWhiteSpace(cncABuscar))
            {
                MostrarNotificacion("Ingrese un codigo CNC valido", Color.Orange);
                return;
            }

            if (isLoading) return;

            await semaphore.WaitAsync();
            try
            {
                isLoading = true;
                MostrarCargando(true, $"Buscando CNC: {cncABuscar}...");

                string queryNestings = @"
                SELECT NstRef, Name as Nombre, Quantity as Cantidad_Programada, 
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
                        NstRef = row["NstRef"].ToString() ?? "",
                        Nombre = row["Nombre"].ToString() ?? "",
                        CantidadProgramada = Convert.ToInt32(row["Cantidad_Programada"]),
                        CantidadReportada = Convert.ToInt32(row["Cantidad_Reportada"]),
                        CantidadPendiente = Convert.ToInt32(row["Cantidad_Pendiente"]),
                        Estado = row["Estado"].ToString() ?? ""
                    });
                }

                cncSeleccionado = cncABuscar;
                lblCNCSeleccionado.Text = $"CNC: {cncSeleccionado}";
                CrearBotonesNestings();

                MostrarNotificacion(nestingsActuales.Count == 0
                    ? $"No se encontraron nestings para el CNC: {cncABuscar}"
                    : $"Se encontraron {nestingsActuales.Count} nestings",
                    nestingsActuales.Count == 0 ? Color.Orange : Color.Green);
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

        private async Task CargarPartesPorNestingAsync(string nstRef, string nestingNombre)
        {
            if (isLoading) return;

            await semaphore.WaitAsync();
            try
            {
                isLoading = true;
                MostrarCargando(true, "Cargando partes...");

                string query = @"
                SELECT MnORef, OprID, PrdRefDst, Quantity
                FROM [MartinRea MJ].[dbo].[DIS_NEST_NEST_00000500]
                WHERE NstRef = @nstRef AND PrdRefDst IS NOT NULL AND PrdRefDst != ''
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
                    partesActuales.Add(new ParteInfo
                    {
                        MnORef = row["MnORef"].ToString() ?? "",
                        OprID = row["OprID"].ToString() ?? "",
                        PrdRefDst = row["PrdRefDst"].ToString() ?? "",
                        Cantidad = Convert.ToInt32(row["Quantity"]),
                        NstRef = nstRef,
                        Recurso = nestingNombre
                    });
                }

                nestingSeleccionado = nestingNombre;
                lblNestingSeleccionado.Text = $"Nesting: {nestingSeleccionado}";
                CrearBotonesPartes();

                MostrarNotificacion(partesActuales.Count == 0
                    ? $"No se encontraron partes"
                    : $"Se encontraron {partesActuales.Count} partes",
                    partesActuales.Count == 0 ? Color.Orange : Color.Green);
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al cargar Partes: {ex.Message}", Color.Red);
            }
            finally
            {
                isLoading = false;
                MostrarCargando(false);
                semaphore.Release();
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

                    lblParteSeleccionada.Text = $"Parte: {parteSeleccionada}";

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
                        MostrarNotificacion($"La parte '{parteSeleccionada}' no tiene CAT_ID asignado en capability", Color.Orange);
                    }
                }
            }
            catch (Exception ex)
            {
                MostrarNotificacion($"Error al obtener informacion: {ex.Message}", Color.Red);
            }
        }

        private void CrearBotonesNestings()
        {
            panelNestings.SuspendLayout();
            panelNestings.Controls.Clear();

            foreach (var nesting in nestingsActuales)
            {
                string estadoTexto = nesting.Estado == "Completado" ? "[Completado]" : (nesting.CantidadPendiente > 0 ? "[Pendiente]" : "[Finalizado]");
                Button btnNesting = new Button
                {
                    Text = $"{estadoTexto} {nesting.Nombre}\nPendiente: {nesting.CantidadPendiente:N0}",
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    BackColor = Color.White,
                    Size = new Size(385, 55),
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand,
                    Margin = new Padding(3),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Tag = nesting,
                    UseVisualStyleBackColor = true
                };
                btnNesting.FlatAppearance.BorderSize = 1;
                btnNesting.FlatAppearance.BorderColor = colorBorde;
                btnNesting.Click += BtnNesting_Click;
                panelNestings.Controls.Add(btnNesting);
            }
            panelNestings.ResumeLayout();
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
                    Text = $"{index:00}. {parte.PrdRefDst}\nCantidad: {parte.Cantidad:N0}",
                    Font = new Font("Arial", 9, FontStyle.Bold),
                    BackColor = Color.White,
                    Size = new Size(385, 55),
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

        private async void BtnNesting_Click(object sender, EventArgs e)
        {
            if (sender is Button btn && btn.Tag is NestingInfo nesting)
            {
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

                LimpiarSeleccionParte();
                await CargarPartesPorNestingAsync(nesting.NstRef, nesting.Nombre);
            }
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

        private void LimpiarSeleccionParte()
        {
            lblParteSeleccionada.Text = "Parte: --";
            lblRecursoSeleccionado.Text = "Recurso: --";
            if (lblCatIdValor != null) lblCatIdValor.Text = "--";
            lblNstRefValor.Text = "NstRef: --";
            lblCantidadInfo.Text = "Programado: 0 | Reportado: 0 | Pendiente: 0";
            parteSeleccionada = "";
            recursoSeleccionado = "";
            catIdSeleccionado = "0";
            nstRefActual = "";
            txtCantidad.Text = "0";
            btnReportar.Enabled = false;

            Control[] controls = panelPrincipal.Controls.Find("lblJobValor", true);
            if (controls.Length > 0 && controls[0] is Label lblJob)
                lblJob.Text = "--";

            panelPartes.Controls.Clear();
        }

        private int ObtenerCantidadIngresada()
        {
            if (string.IsNullOrEmpty(txtCantidad.Text)) return 0;
            return int.TryParse(txtCantidad.Text.Replace(",", ""), out int resultado) ? resultado : 0;
        }

        // ============== METODO PRINCIPAL: REPORTAR PRODUCCION ==============
        // Guarda EN AMBAS TABLAS:
        // 1. TBL_MES_MARS_lASER (tabla original del laser)
        // 2. MES.TBL_PRODUCTION (historico hora por hora)
        private async Task ReportarProduccionAsync()
        {
            if (string.IsNullOrEmpty(nstRefActual))
            {
                MostrarNotificacion("No se encontro informacion para esta parte", Color.Red);
                return;
            }

            if (catIdSeleccionado == "0" || string.IsNullOrEmpty(catIdSeleccionado))
            {
                MostrarNotificacion("La parte seleccionada no tiene un CAT_ID valido en capability", Color.Red);
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
                $"Nesting: {nestingSeleccionado}\n" +
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

                    // 1. Guardar en tabla original del laser (TBL_MES_MARS_lASER)
                    await GuardarEnTablaLaserAsync(cantidadReportar);

                    // 2. Obtener SUBRESOURCE_ID para MES.TBL_PRODUCTION
                    int subresourceId = await ObtenerSubresourceIdAsync(recursoSeleccionado);

                    // 3. Insertar en MES.TBL_PRODUCTION (historico hora por hora)
                    int catIdInt = int.Parse(catIdSeleccionado);
                    await InsertarEnTBLProductionAsync(subresourceId, catIdInt, cantidadReportar);

                    MostrarNotificacion($"Reporte exitoso: {cantidadReportar:N0} piezas\nGuardado en TBL_MES_MARS_lASER\nGuardado en MES.TBL_PRODUCTION", Color.Green);
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

        private async Task RefrescarDatosAsync()
        {
            if (!string.IsNullOrEmpty(cncSeleccionado))
                await BuscarCNCAsync();
            else
                MostrarNotificacion("Primero busque un CNC", Color.Orange);
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
            catIdSeleccionado = "0";
            nstRefActual = "";
            mnORefActual = "";
            cantidadProduccionActual = 0;

            lblNestingSeleccionado.Text = "Nesting: --";
            lblParteSeleccionada.Text = "Parte: --";
            lblRecursoSeleccionado.Text = "Recurso: --";
            if (lblCatIdValor != null) lblCatIdValor.Text = "--";
            lblNstRefValor.Text = "NstRef: --";
            lblCantidadInfo.Text = "Programado: 0 | Reportado: 0 | Pendiente: 0";
            txtCantidad.Text = "0";
            btnReportar.Enabled = false;

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