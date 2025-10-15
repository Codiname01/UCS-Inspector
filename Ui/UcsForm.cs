using Inventor;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Web.Script.Serialization;    // <-- necesario para el Clipboard JSON
using System.Windows.Forms;
// Alias seguros
using DrawColor = System.Drawing.Color;
using SysDir = System.IO.Directory;
using SysEnv = System.Environment;
// Evitar choques con Inventor.*

using SysFile = System.IO.File;
using SysPath = System.IO.Path;
using WinButton = System.Windows.Forms.Button;
using WinComboBox = System.Windows.Forms.ComboBox;
using WinFlow = System.Windows.Forms.FlowLayoutPanel;
using WinLabel = System.Windows.Forms.Label;
// Aliases GUI (evitan ambigüedad con Inventor)
using WinPanel = System.Windows.Forms.Panel;
using WinTable = System.Windows.Forms.TableLayoutPanel;
using WinTextBox = System.Windows.Forms.TextBox;

using SysWin = System.Windows.Forms;
using SysDraw = System.Drawing;





namespace UcsInspectorperu.Ui
{
    public class UcsForm : Form
    {
        private readonly Inventor.Application _app;
   
        private Label lblXo, lblYo, lblZo, lblXa, lblYa, lblZa;
      
      
    
        private (double xo, double yo, double zo, double xa, double ya, double za)? _clipboard;

        // === Recientes (ÚNICO bloque) ===
        // === Recientes (ÚNICO bloque) ===
        // Contenedores

        // Contenedores (campos de UcsForm)
               // ← Panel con AutoScroll
          // ← TableLayoutPanel
    


        private string[] _displayNames = new string[0];
        private System.Windows.Forms.ContextMenuStrip _recentMenu;
        private int _recentMenuIndex = -1;

        // MRU / Favoritos en memoria
        private List<string> _mru = new List<string>();
        private List<string> _fav = new List<string>();


        // Layout
     
       
        private System.Windows.Forms.Panel _scroll;          // contenedor con scroll
     

        // contenedores
     

        // umbral de “modo compacto” (96 DPI base). Puedes mover a settings si quieres.
        private const int COMPACT_PX_AT_96DPI = 860;



        private WinPanel _content;      // Panel con scroll
        private WinTable _header;       // TableLayoutPanel (cabecera)
        private WinTable _grid;         // TableLayoutPanel (cuerpo)
        private WinFlow _recents;      // FlowLayoutPanel (botones recientes)



      
        private int _gridRow = 0; // fila actual en el grid

        // Hints (son LABELS, no TextBox)
        private System.Windows.Forms.Label _hintX, _hintY, _hintZ, _hintRX, _hintRY, _hintRZ;


        // UI principales
        private System.Windows.Forms.ComboBox cboUcs;
        private WinTextBox txtXo, txtYo, txtZo, txtXa, txtYa, txtZa;
        private System.Windows.Forms.TextBox txtFilter; // si no usas alias para filter

        // Botones/acciones
        private System.Windows.Forms.Button btnCenter, btnCopy, btnPaste, btnNext;
        private System.Windows.Forms.Button btnApply, btnRefresh, btnHelp, btnPick;

        // Nudge buttons
        private System.Windows.Forms.Button btnXm, btnXp, btnYm, btnYp, btnZm, btnZp;
        private System.Windows.Forms.Button btnRxm, btnRxp, btnRym, btnRyp, btnRzm, btnRzp;

        // Otros
        private System.Windows.Forms.CheckBox chkDelta;
        private System.Windows.Forms.NumericUpDown numStepMm, numStepDeg;
        private System.Windows.Forms.ToolTip _tt;


        // Aliases GUI (evitan ambigüedad con Inventor)

        private void btnMeasureMm_Click(object sender, EventArgs e) { MeasureDistanceToStepMm(); }
        private void btnMeasureDeg_Click(object sender, EventArgs e) { MeasureAngleToStepDeg(); }

        // Recientes (nuevo layout)

        // Estado
        private readonly List<Inventor.UserCoordinateSystem> _allUcs = new List<Inventor.UserCoordinateSystem>();
        private Inventor.UserCoordinateSystem _ucs;
        private Inventor.UnitsOfMeasure _uom;

        // ⚠️ Elimina estos si aún existen: Panel _recents, List<Button> _recentBtns, Timer _resizeTmr, ReflowHeader()


        public UcsForm(Inventor.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));

            // ── Ventana ───────────────────────────────────────────────────────────────
            Text = "UCS Inspector – Camarada Escate";
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimumSize = new Size(640, 360);
            MaximizeBox = true;
            KeyPreview = true;
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);

            // ── Settings ──────────────────────────────────────────────────────────────
            _cfg = UiSettings.Load() ?? new UiSettings();
            if (_cfg.RecentMax < 6) _cfg.RecentMax = 6;

            // ── UI ────────────────────────────────────────────────────────────────────
            BuildLayout();                 // crea y coloca los controles
            this.AcceptButton = btnApply;  // Enter = Aplicar (sin usar mouse)

            // ToolTip compartido
            _tt = new ToolTip { AutoPopDelay = 5000, InitialDelay = 300, ReshowDelay = 100, ShowAlways = true };

            // ── Estado inicial desde settings ─────────────────────────────────────────
            ApplySettingsToUi(_cfg);       // pasos, pegar Δ, tamaño/posición

          

            // ── Datos ─────────────────────────────────────────────────────────────────
            LoadUcsList();                 // llena el combo
            RefreshRecentButtons();        // MRU + favoritos

            // ── Atajos ────────────────────────────────────────────────────────────────
            WireHotkeysToChildren(this);   // atajos locales
            this.Shown += (s, e) => RegisterGlobalHotkeysIfAny(); // globales si están habilitados en JSON

            // ── Reflow / DPI / tamaño ────────────────────────────────────────────────
            this.Shown += (s, e) => UpdateLayoutForWidth();
            this.Resize += (s, e) => UpdateLayoutForWidth();

            // ── Enter aplica por-caja (tu método existente) ──────────────────────────
            HookEnterEdits();

            // ── Guardado SÓLO al salir o con Enter (evita el bug 45→54) ──────────────
            if (numStepMm != null) { numStepMm.Leave += (s, e) => SaveSettingsFromUi(); numStepMm.KeyDown += OnSaveOnEnter; }
            if (numStepDeg != null) { numStepDeg.Leave += (s, e) => SaveSettingsFromUi(); numStepDeg.KeyDown += OnSaveOnEnter; }
            if (chkDelta != null) { chkDelta.CheckedChanged += (s, e) => SaveSettingsFromUi(); }

            // ── Limpieza ──────────────────────────────────────────────────────────────
            this.FormClosed += (s, e) =>
            {
                try { GlobalHotkeys.UnregisterAll(); } catch { }
                try { _tt?.Dispose(); } catch { }
                _tt = null;
            };
        }




        private void OnSaveOnEnter(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SaveSettingsFromUi(); // solo guarda; NO recarga la UI aquí
                e.Handled = true;
            }
        }



        // Redondea y respeta límites del NumericUpDown
        private static decimal ClampToNumeric(NumericUpDown nud, decimal v)
        {
            if (nud == null) return v;
            if (v < nud.Minimum) v = nud.Minimum;
            if (v > nud.Maximum) v = nud.Maximum;
            return decimal.Round(v, nud.DecimalPlaces);
        }

        // Setters centralizados + guardado
        private void SetStepMm(double mm)
        {
            if (numStepMm != null)
            {
                numStepMm.Value = ClampToNumeric(numStepMm, (decimal)mm);
                SaveSettingsFromUi(); // solo guarda, NO recarga UI
            }
        }
        private void SetStepDeg(double deg)
        {
            if (numStepDeg != null)
            {
                // Útil: normaliza a [-360, 360]
                while (deg > 360.0) deg -= 360.0;
                while (deg < -360.0) deg += 360.0;

                numStepDeg.Value = ClampToNumeric(numStepDeg, (decimal)deg);
                SaveSettingsFromUi();
            }
        }


        // Distancia mínima entre dos picks (mm) y la pone en numStepMm
        // Distancia mínima entre dos picks (mm) y la pone en numStepMm
        private void MeasureDistanceToStepMm()
        {
            try
            {
                var cm = _app.CommandManager;
                object a = cm.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Selecciona el primer elemento");
                object b = cm.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Selecciona el segundo elemento");

                // Devuelve distancia en unidades de base (cm)
                double dDb = _app.MeasureTools.GetMinimumDistance(a, b);

                // Convertir cm (DB) → mm usando el UnitsOfMeasure del documento
                double mm = dDb;
                var doc = _app.ActiveDocument as Inventor.Document;
                if (doc != null && doc.UnitsOfMeasure != null)
                {
                    mm = doc.UnitsOfMeasure.ConvertUnits(
                        dDb, UnitsTypeEnum.kDatabaseLengthUnits, UnitsTypeEnum.kMillimeterLengthUnits);
                }
                else
                {
                    // Fallback seguro (cm → mm)
                    mm = dDb * 10.0;
                }

                SetStepMm(mm); // tu setter que guarda sin recargar la UI
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Pick cancelado / E_FAIL: ignorar
            }
            catch { /* opcional: log */ }
        }

        // Ángulo entre dos picks (grados) y lo pone en numStepDeg (esto sí estaba ok)
        private void MeasureAngleToStepDeg()
        {
            try
            {
                var cm = _app.CommandManager;
                object a = cm.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Elemento A (cara/eje/arista)");
                object b = cm.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Elemento B (cara/eje/arista)");

                double rad = _app.MeasureTools.GetAngle(a, b); // radianes
                double deg = rad * 180.0 / Math.PI;

                SetStepDeg(deg);
            }
            catch (System.Runtime.InteropServices.COMException) { }
            catch { }
        }






        private void HookEnterEdits()
        {
            if (txtXo != null) txtXo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { ApplyOneFromBox(txtXo, false); e.Handled = true; } };
            if (txtYo != null) txtYo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { ApplyOneFromBox(txtYo, false); e.Handled = true; } };
            if (txtZo != null) txtZo.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { ApplyOneFromBox(txtZo, false); e.Handled = true; } };
            if (txtXa != null) txtXa.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { ApplyOneFromBox(txtXa, true); e.Handled = true; } };
            if (txtYa != null) txtYa.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { ApplyOneFromBox(txtYa, true); e.Handled = true; } };
            if (txtZa != null) txtZa.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { ApplyOneFromBox(txtZa, true); e.Handled = true; } };
        }


        private void ReflowHeader()
        {
            if (btnPick == null || cboUcs == null || txtFilter == null || _recents == null) return;

            btnPick.Left = ClientSize.Width - btnPick.Width - 12;
            cboUcs.Left = txtFilter.Right + 8;

            int rightSpace = btnPick.Left - 8;
            int newWidth = rightSpace - cboUcs.Left;
            if (newWidth < 160) newWidth = 160;
            cboUcs.Width = newWidth;

            _recents.Width = ClientSize.Width - 24;
        }

        private void RegisterGlobalHotkeysIfAny()
        {
            try
            {
                // Asegura config
                if (_cfg == null)
                    _cfg = UiSettings.Load() ?? new UiSettings();

                // Limpia todo antes de registrar
                GlobalHotkeys.UnregisterAll();

                if (_cfg == null || !_cfg.GlobalNudgeHotkeys)
                    return;

                var list = new List<Tuple<uint, uint, Action>>();

                // Recientes (1..6)
                var rh = _cfg.RecentHotkeys;
                int max = (rh != null) ? Math.Min(6, rh.Length) : 0;
                for (int i = 0; i < max; i++)
                {
                    HotkeyUtil.Parsed p;
                    if (!string.IsNullOrEmpty(rh[i]) &&
                        HotkeyUtil.TryParse(rh[i], out p) && p.IsValid)
                    {
                        int captured = i;
                        list.Add(Tuple.Create(p.FsModifiers, p.VirtualKey,
                            (Action)(() => SelectRecentByIndex(captured))));
                    }
                }

                // Pick en pantalla
                HotkeyUtil.Parsed pp;
                if (!string.IsNullOrEmpty(_cfg.PickHotkey) &&
                    HotkeyUtil.TryParse(_cfg.PickHotkey, out pp) && pp.IsValid)
                {
                    list.Add(Tuple.Create(pp.FsModifiers, pp.VirtualKey, (Action)TryPickUcs));
                }

                // Registrar si hay algo y _app está disponible
                if (list.Count > 0 && _app != null)
                    GlobalHotkeys.RegisterMany(new IntPtr(_app.MainFrameHWND), list);
            }
            catch { /* opcional: log */ }
        }




        // ======= Settings del UI (persistencia en JSON) =======
        [DataContract]
      
        public class UiSettings
        {
            // === UI / comportamiento ===
            [DataMember(Order = 1)] public double StepMm { get; set; } = 1.0;
            [DataMember(Order = 2)] public double StepDeg { get; set; } = 1.0;
            [DataMember(Order = 3)] public bool PasteAsDelta { get; set; } = false;

            [DataMember(Order = 4)] public int Left { get; set; } = 32;
            [DataMember(Order = 5)] public int Top { get; set; } = 32;
            [DataMember(Order = 6)] public int Width { get; set; } = 780;
            [DataMember(Order = 7)] public int Height { get; set; } = 480;

            [DataMember(Order = 8)] public bool StartWithLastUcs { get; set; } = true;

            [DataMember(Order = 9)] public int RecentMax { get; set; } = 6;   // mínimo 6
            [DataMember(Order = 10)] public int RecentCapacity { get; set; } = 24;
            [DataMember(Order = 11)] public string[] RecentUcs { get; set; } = new string[0];
            [DataMember(Order = 12)] public string[] FavoriteUcs { get; set; } = new string[0];

            // Atajos configurables
            [DataMember(Order = 13)]
            public string[] RecentHotkeys { get; set; } =
                new[] { "Alt+1", "Alt+2", "Alt+3", "Alt+4", "Alt+5", "Alt+6" };

            [DataMember(Order = 14)] public string PickHotkey { get; set; } = "Alt+P";

            // Globales (si los usas)
            [DataMember(Order = 15)] public bool GlobalNudgeHotkeys { get; set; } = false;
            [DataMember(Order = 16)] public bool UseGlobalHotkey { get; set; } = true;
            [DataMember(Order = 17)] public string GlobalHotkey { get; set; } = "Ctrl+Shift+U";

            // Campos legados que alguna vez aparecieron (para no romper carga)
            [DataMember] public int CompactPxAt96Dpi { get; set; } = 860;
            [DataMember] public string[] Favorites { get; set; } = new string[0];

            // UiSettings.cs
            [DataMember(Order = 20)]
            public string ToggleHotkey { get; set; } = "Alt+F12";
            [DataMember(Order = 30)]
            public bool HintsAsTooltips { get; set; } = true;   // por defecto: usar ToolTips







            // === RUTA === (igual que ya tenías: %LOCALAPPDATA%\UcsInspectorperu\settings.json)
            private static string GetPath()
            {
                string dir = SysPath.Combine(
                    SysEnv.GetFolderPath(SysEnv.SpecialFolder.LocalApplicationData),
                    "UcsInspectorperu");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                return SysPath.Combine(dir, "settings.json");
            }

            // Normaliza valores nulos/incorrectos
            private void EnsureDefaults()
            {
                if (RecentUcs == null) RecentUcs = new string[0];
                if (FavoriteUcs == null) FavoriteUcs = new string[0];
                if (RecentHotkeys == null || RecentHotkeys.Length == 0)
                    RecentHotkeys = new[] { "Alt+1", "Alt+2", "Alt+3", "Alt+4", "Alt+5", "Alt+6" };

                if (RecentMax < 6) RecentMax = 6;

                if (Width < 400) Width = 780;
                if (Height < 300) Height = 480;

                // Si existía el campo antiguo "Favorites", fusiónalo
                if (Favorites != null && Favorites.Length > 0)
                {
                    var set = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var f in FavoriteUcs) set.Add(f ?? "");
                    foreach (var f in Favorites) if (!string.IsNullOrWhiteSpace(f)) set.Add(f);
                    FavoriteUcs = new System.Collections.Generic.List<string>(set).ToArray();
                    Favorites = new string[0];
                }
            }

            public static UiSettings Load()
            {
                try
                {
                    var p = GetPath();
                    if (SysFile.Exists(p))
                    {
                        using (var fs = SysFile.OpenRead(p))
                        {
                            var ser = new DataContractJsonSerializer(typeof(UiSettings));
                            var s = (UiSettings)ser.ReadObject(fs);
                            if (s == null) s = new UiSettings();
                            s.EnsureDefaults();
                            return s;
                        }
                    }
                }
                catch { /* ignorar y devolver defaults */ }

                // Primera ejecución: crear archivo con defaults
                var def = new UiSettings();
                def.EnsureDefaults();
                def.Save();
                return def;
            }

            public void Save()
            {
                try
                {
                    var p = GetPath();
                    using (var fs = SysFile.Create(p))
                    {
                        var ser = new DataContractJsonSerializer(typeof(UiSettings));
                        ser.WriteObject(fs, this);
                    }
                }
                catch { /* ignorar errores de IO */ }
            }
        }

        // campo único (evita ambigüedad CS0229)
        private UiSettings _cfg;


        private string SettingsPath()
        {
            string dir = SysPath.Combine(
                SysEnv.GetFolderPath(SysEnv.SpecialFolder.LocalApplicationData),
                "UcsInspectorperu");
            try { SysDir.CreateDirectory(dir); } catch { }
            return SysPath.Combine(dir, "settings.json");
        }

        private void LoadSettings()
        {
            try
            {
                var p = SettingsPath();
                if (!SysFile.Exists(p)) return;

                var ser = new DataContractJsonSerializer(typeof(UiSettings));
                using (var fs = SysFile.OpenRead(p))
                    _cfg = (UiSettings)ser.ReadObject(fs);

                if (_cfg.RecentUcs == null) _cfg.RecentUcs = new string[0];
                if (_cfg.RecentMax <= 0) _cfg.RecentMax = 3;
            }
            catch { /* usa defaults si falla */ }
        }

        private void SaveSettingsFromUi()
        {
            if (_cfg == null) _cfg = new UiSettings();

            _cfg.StepMm = (double)numStepMm.Value;
            _cfg.StepDeg = (double)numStepDeg.Value;
            _cfg.PasteAsDelta = chkDelta.Checked;

            _cfg.Left = this.Left;
            _cfg.Top = this.Top;
            _cfg.Width = this.Width;
            _cfg.Height = this.Height;

            _cfg.Save();          // ✅ en lugar de SaveSettings();
        }

        private void TrackRecentUcs(string name)
        {
            if (string.IsNullOrWhiteSpace(name) || _cfg == null) return;

            var list = new List<string>(_cfg.RecentUcs ?? new string[0]);
            list.RemoveAll(s => string.Equals(s, name, StringComparison.OrdinalIgnoreCase));
            list.Insert(0, name);

            int max = _cfg.RecentMax > 0 ? _cfg.RecentMax : 3;
            while (list.Count > max) list.RemoveAt(list.Count - 1);

            _cfg.RecentUcs = list.ToArray();
            _cfg.Save();                    // ✅ nuevo
            RefreshRecentButtons();
        }

        // 1) Favoritos (todos) + MRU (recortado a RecentMax, min 6). Sin duplicados.
        private string[] BuildPinnedRecentNames()
        {
            var fav = _cfg?.Favorites ?? Array.Empty<string>();
            var mru = _cfg?.RecentUcs ?? Array.Empty<string>();

            int capMru = (_cfg != null && _cfg.RecentMax > 0) ? _cfg.RecentMax : 6;

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var list = new List<string>();

            // Favoritos siempre (en su orden)
            foreach (var n in fav)
            {
                if (string.IsNullOrWhiteSpace(n)) continue;
                if (seen.Add(n)) list.Add(n);
            }

            // MRU (excluye favoritos), limitado
            foreach (var n in mru)
            {
                if (string.IsNullOrWhiteSpace(n)) continue;
                if (seen.Contains(n)) continue;
                if (seen.Add(n)) list.Add(n);
                if (--capMru <= 0) break;
            }

            return list.ToArray();
        }

        // 2) Tooltip rico (opcional)
        private string BuildRecentTooltip(string name, bool exists, bool isFav, string gesture)
        {
            if (!exists) return "Este UCS no existe en el documento actual";
            var extra = new List<string>();
            if (!string.IsNullOrEmpty(gesture)) extra.Add(gesture);
            if (isFav) extra.Add("favorito");
            return "Seleccionar " + name + (extra.Count > 0 ? " — " + string.Join(" · ", extra) : "");
        }

        // 3) Dibuja los botones de recientes con favoritos “pineados” a la izquierda
        private void RefreshRecentButtons()
        {
            if (_recents == null || _cfg == null) return;

            var tt = EnsureTooltip();
            EnsureRecentMenu();

            _displayNames = BuildPinnedRecentNames();
            if (_displayNames == null || _displayNames.Length == 0)
            {
                _recents.Controls.Clear();
                return;
            }

            _recents.SuspendLayout();
            try
            {
                _recents.Controls.Clear();

                for (int i = 0; i < _displayNames.Length; i++)
                {
                    string name = _displayNames[i];
                    bool fav = IsFavorite(name);

                    string gesture = (_cfg.RecentHotkeys != null && i < _cfg.RecentHotkeys.Length)
                                     ? _cfg.RecentHotkeys[i]
                                     : null;

                    string caption = (fav ? "★ " : "")
                                   + (string.IsNullOrEmpty(gesture) ? "" : "[" + gesture + "] ")
                                   + name;

                    var btn = new System.Windows.Forms.Button
                    {
                        AutoSize = true,
                        AutoEllipsis = true,
                        Height = 24,
                        Text = caption,
                        Tag = i,
                        Margin = new Padding(0, 0, 6, 0),
                        UseVisualStyleBackColor = true
                    };

                    bool exists = (FindIndexByName(name) >= 0);
                    btn.Enabled = exists;

                    tt.SetToolTip(btn, BuildRecentTooltip(name, exists, fav, gesture));

                    int capturedIndex = i;
                    string capturedName = name;

                    btn.Click += (s, e) =>
                    {
                        if (!SelectRecentByIndex(capturedIndex))
                        {
                            MessageBox.Show(
                                "El UCS '" + capturedName + "' no existe en este documento.",
                                "UCS Inspector",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    };

                    btn.MouseUp += (s, e) =>
                    {
                        if (e.Button != MouseButtons.Right) return;
                        _recentMenuIndex = capturedIndex;

                        // 0 = Marcar favorito, 1 = Quitar favorito, 2 = sep, 3 = Quitar de recientes
                        _recentMenu.Items[0].Visible = !IsFavorite(capturedName);
                        _recentMenu.Items[1].Visible = IsFavorite(capturedName);

                        _recentMenu.Show(btn, e.Location);
                    };

                    _recents.Controls.Add(btn);
                }
            }
            finally
            {
                _recents.ResumeLayout();
            }
        }

        // UcsForm.cs
        private void SafeToggleShow()
        {
            if (this.IsDisposed) return;

            if (this.InvokeRequired)
            {
                this.BeginInvoke((Action)SafeToggleShow);
                return;
            }

            try
            {
                if (!this.Visible)
                {
                    this.Show();
                    if (this.WindowState == FormWindowState.Minimized)
                        this.WindowState = FormWindowState.Normal;

                    this.Activate();
                    // “pop” al frente sin quedarse topmost
                    this.TopMost = true;
                    this.TopMost = false;
                }
                else
                {
                    this.Hide();
                }
            }
            catch { /* opcional: log */ }
        }


        private void EnsureRecentMenu()
        {
            if (_recentMenu != null) return;

            _recentMenu = new ContextMenuStrip();
            var miFav = new ToolStripMenuItem("Marcar como favorito");
            var miUnfav = new ToolStripMenuItem("Quitar de favoritos");
            var miRemove = new ToolStripMenuItem("Quitar de recientes");

            miFav.Click += (s, e) => ToggleFavorite(true);
            miUnfav.Click += (s, e) => ToggleFavorite(false);
            miRemove.Click += (s, e) => RemoveFromRecent();

            _recentMenu.Items.Add(miFav);
            _recentMenu.Items.Add(miUnfav);
            _recentMenu.Items.Add(new ToolStripSeparator());
            _recentMenu.Items.Add(miRemove);
        }

        private void ToggleFavorite(bool add)
        {
            if (_displayNames == null || _recentMenuIndex < 0 || _recentMenuIndex >= _displayNames.Length) return;
            string name = _displayNames[_recentMenuIndex];
            SetFavorite(name, add);
            _cfg.Save();                 // << antes tenías UiSettings.Save(_cfg) (no existe)
            RefreshRecentButtons();
        }

        private void RemoveFromRecent()
        {
            if (_cfg == null || _displayNames == null || _recentMenuIndex < 0 || _recentMenuIndex >= _displayNames.Length) return;

            string name = _displayNames[_recentMenuIndex];
            var list = new List<string>(_cfg.RecentUcs ?? Array.Empty<string>());
            list.RemoveAll(n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase));
            _cfg.RecentUcs = list.ToArray();
            _cfg.Save();                 // << idem
            RefreshRecentButtons();
        }






        private void AlignViewToUcs(Inventor.UserCoordinateSystem u)
        {
            if (u == null || ActiveDoc == null) return;

            WithFast(() =>
            {
                try
                {
                    var tg = _app.TransientGeometry;
                    var cam = _app.ActiveView.Camera;

                    // --- Origen y ejes del UCS (reflexión para ser robustos) ---
                    Inventor.Point o = null;
                    Inventor.Vector x = null, y = null, z = null;

                    object GetProp(object obj, string name)
                    {
                        try { return obj.GetType().InvokeMember(name, BindingFlags.GetProperty, null, obj, new object[0]); }
                        catch { return null; }
                    }

                    // Origen
                    o = GetProp(u, "Origin") as Inventor.Point;
                    if (o == null)
                    {
                        // Fallback: usar offsets actuales de la UI
                        double ox = GetVal((Inventor.Parameter)txtXo.Tag);
                        double oy = GetVal((Inventor.Parameter)txtYo.Tag);
                        double oz = GetVal((Inventor.Parameter)txtZo.Tag);
                        o = tg.CreatePoint(ox, oy, oz);
                    }

                    // Ejes (distintas propiedades según interop)
                    Inventor.Vector TryVec(object obj, string[] names)
                    {
                        foreach (var n in names)
                        {
                            var v = GetProp(obj, n);
                            if (v is Inventor.Vector) return (Inventor.Vector)v;

                            // Algunas variantes exponen UnitVector → convertir a Vector
                            if (v != null && v.GetType().Name.IndexOf("UnitVector", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                try
                                {
                                    double vx = (double)v.GetType().InvokeMember("X", BindingFlags.GetProperty, null, v, new object[0]);
                                    double vy = (double)v.GetType().InvokeMember("Y", BindingFlags.GetProperty, null, v, new object[0]);
                                    double vz = (double)v.GetType().InvokeMember("Z", BindingFlags.GetProperty, null, v, new object[0]);
                                    return tg.CreateVector(vx, vy, vz);
                                }
                                catch { }
                            }
                        }
                        return null;
                    }

                    x = TryVec(u, new[] { "XAxis", "XDirection", "XAxisVector" });
                    y = TryVec(u, new[] { "YAxis", "YDirection", "YAxisVector" });
                    z = TryVec(u, new[] { "ZAxis", "ZDirection", "ZAxisVector" });

                    // --- Plan B: componer ejes desde RX/RY/RZ si algo faltó ---
                    if (x == null || y == null || z == null)
                    {
                        double rx = GetVal((Inventor.Parameter)txtXa.Tag) * Math.PI / 180.0;
                        double ry = GetVal((Inventor.Parameter)txtYa.Tag) * Math.PI / 180.0;
                        double rz = GetVal((Inventor.Parameter)txtZa.Tag) * Math.PI / 180.0;

                        double ca = Math.Cos(rz), sa = Math.Sin(rz);
                        double cb = Math.Cos(ry), sb = Math.Sin(ry);
                        double cc = Math.Cos(rx), sc = Math.Sin(rx);

                        x = x ?? tg.CreateVector(ca * cb, ca * sb * sc - sa * cc, ca * sb * cc + sa * sc);
                        y = y ?? tg.CreateVector(sa * cb, sa * sb * sc + ca * cc, sa * sb * cc - ca * sc);
                        z = z ?? tg.CreateVector(-sb, cb * sc, cb * cc);
                    }

                    try { x.Normalize(); } catch { }
                    try { y.Normalize(); } catch { }
                    try { z.Normalize(); } catch { }

                    // Distancia: usa la actual si existe; si no, 200
                    double curDist = 200.0;
                    try
                    {
                        var cur = tg.CreateVector(
                            cam.Eye.X - cam.Target.X,
                            cam.Eye.Y - cam.Target.Y,
                            cam.Eye.Z - cam.Target.Z);
                        curDist = Math.Max(10.0, cur.Length);
                    }
                    catch { }

                    // Eye = origen + (+Y)*dist  → se mira hacia -Y del UCS
                    var eye = tg.CreatePoint(
                        o.X + y.X * curDist,
                        o.Y + y.Y * curDist,
                        o.Z + y.Z * curDist);

                    cam.Eye = eye;
                    cam.Target = o;

                    // Up = +Z del UCS (la cámara exige UnitVector)
                    Inventor.UnitVector up;
                    try { up = tg.CreateUnitVector(z.X, z.Y, z.Z); }
                    catch { up = tg.CreateUnitVector(0, 0, 1); } // fallback seguro

                    cam.UpVector = up;
                    cam.Apply();
                    _app.ActiveView.Update();

                    // Encadrar con la orientación ya aplicada
                    try { _app.ActiveView.Fit(true); } catch { }
                }
                catch
                {
                    // silencioso
                }
            });
        }




        // ---------- Clipboard JSON (interoperable) ----------

        [DataContract]
        private class ClipPacket
        {
            // marcador y versión para reconocer nuestro JSON
            [DataMember] public string kind = "UCS-expr";
            [DataMember] public int ver = 2;

            // valores con unidades, p.ej. "92 mm" / "10 deg"
            [DataMember] public string xo, yo, zo;
            [DataMember] public string xa, ya, za;

            // expresiones literales (opcionales); p.ej. "=d0*2", "25 mm"
            [DataMember] public string xoExp, yoExp, zoExp;
            [DataMember] public string xaExp, yaExp, zaExp;
        }

        // JSON helpers (sin System.Web.Extensions)
        private static string ToJson<T>(T obj)
        {
            try
            {
                var ser = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(T));
                using (var ms = new System.IO.MemoryStream())
                {
                    ser.WriteObject(ms, obj);
                    return System.Text.Encoding.UTF8.GetString(ms.ToArray());
                }
            }
            catch { return null; }
        }

        private static bool TryFromJson<T>(string s, out T obj) where T : class
        {
            obj = null;
            try
            {
                var ser = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(T));
                using (var ms = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(s)))
                {
                    obj = ser.ReadObject(ms) as T;
                    return obj != null;
                }
            }
            catch { return false; }
        }






        private void AddNudgePair(int left, int top, string textMinus, string textPlus,
                                  Action onMinus, Action onPlus,
                                  out Button bMinus, out Button bPlus)
        {
            bMinus = new Button { Left = left, Top = top, Width = 60, Height = 24, Text = textMinus };
            bPlus = new Button { Left = left + 65, Top = top, Width = 60, Height = 24, Text = textPlus };
            bMinus.Click += (s, e) => onMinus();
            bPlus.Click += (s, e) => onPlus();
            Controls.Add(bMinus); Controls.Add(bPlus);
        }

        private void MakeRow(string title, int top, out Label lab, out WinTextBox box)
        {
            var lblT = new Label { Left = 12, Top = top, Width = 100, Text = title };
            lab = new Label { Left = 120, Top = top, Width = 220, Text = "—" };

            var tb = new WinTextBox { Left = 350, Top = top - 3, Width = 185 };
            tb.ForeColor = DrawColor.Gray; tb.Text = "ej. +10 mm / *1.1 / 25 mm";
            tb.GotFocus += (s, e) => { if (tb.ForeColor == DrawColor.Gray) { tb.Text = ""; tb.ForeColor = DrawColor.Black; } };
            tb.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(tb.Text)) { tb.ForeColor = DrawColor.Gray; tb.Text = "ej. +10 mm / *1.1 / 25 mm"; } };

            Controls.Add(lblT); Controls.Add(lab); Controls.Add(tb);
            box = tb;
        }

        private Inventor.Document ActiveDoc { get { return _app.ActiveDocument; } }

        private bool _suspendSelChanged; // guardia anti-reentradas

        // Guardia anti-reentradas
      

        // === 1) Cargar lista desde el documento ===
        private void LoadUcsList()
        {
            // limpia
            cboUcs.Items.Clear();
            _allUcs.Clear();
            _uom = (ActiveDoc != null) ? ActiveDoc.UnitsOfMeasure : null;

            // colecciones según tipo
            Inventor.UserCoordinateSystems col = null;
            var part = ActiveDoc as Inventor.PartDocument;
            var asm = ActiveDoc as Inventor.AssemblyDocument;
            if (part != null) col = part.ComponentDefinition.UserCoordinateSystems;
            else if (asm != null) col = asm.ComponentDefinition.UserCoordinateSystems;

            if (col == null || col.Count == 0)
            {
                MessageBox.Show("No hay UCS en el documento actual.");
                return;
            }

            for (int i = 1; i <= col.Count; i++) _allUcs.Add(col[i]);

            // evento: UNA sola vez
            cboUcs.DisplayMember = "Name";
            cboUcs.SelectedIndexChanged -= CboUcs_SelectedIndexChanged;
            cboUcs.SelectedIndexChanged += CboUcs_SelectedIndexChanged;

            // aplica filtro SIN disparar evento...
            _suspendSelChanged = true;
            try { ApplyFilterCore(); } finally { _suspendSelChanged = false; }

            // ...y AHORA fuerza la carga del UCS seleccionado
            LoadCurrentUcsFromCombo();

            RefreshRecentButtons();
        }

        // === 2) Handler único del combo ===
        private void CboUcs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suspendSelChanged) return;

            LoadCurrentUcsFromCombo(); // establece _ucs y carga valores

            if (_ucs != null) { TouchRecent(_ucs.Name); RefreshRecentButtons(); }
        }

        // === 3) Filtrado (núcleo) ===
        // NO llama a SetCurrentUcs aquí porque suele usarse con _suspendSelChanged=true
        private void ApplyFilterCore()
        {
            string f = (txtFilter != null && txtFilter.ForeColor != DrawColor.Gray)
                       ? (txtFilter.Text ?? "").Trim()
                       : "";

            // recuerda selección previa si existe
            string prev = null;
            var prevU = cboUcs.SelectedItem as Inventor.UserCoordinateSystem;
            if (prevU != null) prev = prevU.Name;

            cboUcs.BeginUpdate();
            try
            {
                cboUcs.Items.Clear();

                for (int i = 0; i < _allUcs.Count; i++)
                {
                    var u = _allUcs[i];
                    if (string.IsNullOrEmpty(f) ||
                        u.Name.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        cboUcs.Items.Add(u);
                    }
                }

                // prioridad de selección: anterior → MRU[0] → primero
                int sel = FindIndexByName(prev);
                if (sel < 0 && _cfg != null && _cfg.RecentUcs != null && _cfg.RecentUcs.Length > 0)
                    sel = FindIndexByName(_cfg.RecentUcs[0]);
                if (sel < 0 && cboUcs.Items.Count > 0) sel = 0;

                cboUcs.SelectedIndex = sel; // aquí puede no disparar por el guardia externo
            }
            finally { cboUcs.EndUpdate(); }
        }

        // === helper: carga el UCS del combo y vincula parámetros ===
        private void LoadCurrentUcsFromCombo()
        {
            var u = cboUcs.SelectedItem as Inventor.UserCoordinateSystem;
            if (u != null)
            {
                SetCurrentUcs(u);      // <-- aquí asignas _ucs, pones txtXo.Tag,... y llamas a LoadValues()
            }
            else
            {
                _ucs = null;
                ClearUcsUi();          // si tienes este método para poner “—”
            }
        }

        // busca por nombre en el combo
        private int FindIndexByName(string name)
        {
            if (string.IsNullOrEmpty(name)) return -1;
            for (int i = 0; i < cboUcs.Items.Count; i++)
            {
                var u = (Inventor.UserCoordinateSystem)cboUcs.Items[i];
                if (string.Equals(u.Name, name, StringComparison.OrdinalIgnoreCase)) return i;
            }
            return -1;
        }
        // Wrapper para mantener compatibilidad con el código viejo.
        // Apunta al núcleo nuevo ApplyFilterCore().
        private void ApplyFilter()
        {
            ApplyFilterCore();
        }
        // Establece el UCS actual y recarga los parámetros/valores en la UI.
        // Si tu LoadValues() ya asigna txtXo.Tag, txtYo.Tag, ... no necesitas más.
        private void SetCurrentUcs(Inventor.UserCoordinateSystem u)
        {
            _ucs = u;
            try
            {
                // Aquí tu lógica existente que pinta los campos:
                // - asignar .Tag de cada textbox a su Inventor.Parameter
                // - leer valores del UCS y mostrarlos
                LoadValues();
            }
            catch { /* opcional: Log("SetCurrentUcs/LoadValues error: " + ex.Message); */ }
        }
        // Limpia textos y elimina vínculos a parámetros (Tag = null)
        private void ClearUcsUi()
        {
            const string dash = "—";
            Action<WinTextBox> wipe = tb =>
            {
                if (tb == null) return;
                tb.Text = dash;
                tb.Tag = null; // muy importante: los nudges no deben usar Tags viejos
            };

            wipe(txtXo); wipe(txtYo); wipe(txtZo);
            wipe(txtXa); wipe(txtYa); wipe(txtZa);
        }



        private void TryPickUcs()
        {
            try
            {
                object picked = _app.CommandManager.Pick(
                    Inventor.SelectionFilterEnum.kAllEntitiesFilter, "Selecciona un UCS…");

                if (picked is Inventor.UserCoordinateSystem)
                { SelectUcs((Inventor.UserCoordinateSystem)picked); return; }

                var t = picked != null ? picked.GetType() : null;
                if (t != null && t.Name.Contains("UserCoordinateSystem"))
                { SelectUcs((Inventor.UserCoordinateSystem)picked); return; }

                MessageBox.Show("El objeto seleccionado no es un UCS. Usa la lista.");
            }
            catch { /* cancelado */ }
        }


        // ¿El parámetro es de solo lectura? (vía reflexión; compatible con distintos interop)
        private bool IsReadOnly(Inventor.Parameter p)
        {
            try
            {
                if (p == null) return true;
                var prop = p.GetType().GetProperty("ReadOnly");
                if (prop != null)
                {
                    object v = prop.GetValue(p, null);
                    if (v is bool) return (bool)v;
                }
            }
            catch { }

            // Extra: si el tipo es ReferenceParameter, trátalo como RO
            try
            {
                var tprop = p.GetType().GetProperty("ParameterType");
                if (tprop != null)
                {
                    var tval = tprop.GetValue(p, null);
                    if (tval != null &&
                        tval.ToString().IndexOf("kReferenceParameter", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }
            catch { }

            return false;
        }



        private bool SelectRecentByIndex(int idx)
        {
            if (_cfg == null || _cfg.RecentMax <= 0) return false;
            if (_cfg.RecentUcs == null || idx < 0 || idx >= _cfg.RecentUcs.Length) return false;

            string name = _cfg.RecentUcs[idx];
            for (int i = 0; i < cboUcs.Items.Count; i++)
            {
                var u = cboUcs.Items[i] as Inventor.UserCoordinateSystem;
                if (u != null && string.Equals(u.Name, name, StringComparison.OrdinalIgnoreCase))
                { cboUcs.SelectedIndex = i; CenterOnUcs(u); return true; }
            }
            System.Media.SystemSounds.Beep.Play();
            return false;
        }


        // Habilita/Deshabilita la fila (textbox y dos botones nudge)
        private void SetRowEnabled(WinTextBox tb, Button bMinus, Button bPlus, bool enabled)
        {
            if (tb != null)
            {
                tb.Enabled = enabled;
                tb.BackColor = enabled ? System.Drawing.SystemColors.Window
                                       : System.Drawing.SystemColors.ControlLight;
            }
            if (bMinus != null) bMinus.Enabled = enabled;
            if (bPlus != null) bPlus.Enabled = enabled;
        }


        private void SelectUcs(Inventor.UserCoordinateSystem u)
        {
            for (int i = 0; i < cboUcs.Items.Count; i++)
                if (ReferenceEquals(cboUcs.Items[i], u)) { cboUcs.SelectedIndex = i; return; }
            _ucs = u; LoadValues();
        }

        private void LoadValues()
        {
            if (_ucs == null) return;

            var px = TryGetParam(_ucs, "XOffset") ?? FindParamByName(_ucs, "X", false);
            var py = TryGetParam(_ucs, "YOffset") ?? FindParamByName(_ucs, "Y", false);
            var pz = TryGetParam(_ucs, "ZOffset") ?? FindParamByName(_ucs, "Z", false);
            var ax = TryGetParam(_ucs, "XAngle") ?? FindParamByName(_ucs, "X", true);
            var ay = TryGetParam(_ucs, "YAngle") ?? FindParamByName(_ucs, "Y", true);
            var az = TryGetParam(_ucs, "ZAngle") ?? FindParamByName(_ucs, "Z", true);

            lblXo.Text = ToStr(px); lblYo.Text = ToStr(py); lblZo.Text = ToStr(pz);
            lblXa.Text = ToStr(ax); lblYa.Text = ToStr(ay); lblZa.Text = ToStr(az);

            txtXo.Tag = px; txtYo.Tag = py; txtZo.Tag = pz;
            txtXa.Tag = ax; txtYa.Tag = ay; txtZa.Tag = az;

            // Bloquea edición si el parámetro es de solo lectura
            SetRowEnabled(txtXo, btnXm, btnXp, !IsReadOnly(px));
            SetRowEnabled(txtYo, btnYm, btnYp, !IsReadOnly(py));
            SetRowEnabled(txtZo, btnZm, btnZp, !IsReadOnly(pz));
            SetRowEnabled(txtXa, btnRxm, btnRxp, !IsReadOnly(ax));
            SetRowEnabled(txtYa, btnRym, btnRyp, !IsReadOnly(ay));
            SetRowEnabled(txtZa, btnRzm, btnRzp, !IsReadOnly(az));
        }


        private string ToStr(Inventor.Parameter p)
        {
            if (p == null) return "—";
            try { return p.Expression; }
            catch { return Convert.ToString(p.Value, CultureInfo.InvariantCulture); }
        }

        private void ApplyChanges()
        {
            if (_ucs == null) return;
            WithFast(() =>
            {
                var tm = _app.TransactionManager; Inventor.Transaction tx = null;
                try
                {
                    tx = tm.StartTransaction((Inventor._Document)ActiveDoc, "UCS apply");

                    ApplyOne((Inventor.Parameter)txtXo.Tag, txtXo.Text, false);
                    ApplyOne((Inventor.Parameter)txtYo.Tag, txtYo.Text, false);
                    ApplyOne((Inventor.Parameter)txtZo.Tag, txtZo.Text, false);
                    ApplyOne((Inventor.Parameter)txtXa.Tag, txtXa.Text, true);
                    ApplyOne((Inventor.Parameter)txtYa.Tag, txtYa.Text, true);
                    ApplyOne((Inventor.Parameter)txtZa.Tag, txtZa.Text, true);

                    ActiveDoc.Update(); tx.End(); LoadValues();
                    MessageBox.Show("Cambios aplicados.");
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    if (tx != null) tx.Abort();
                    if (ex.ErrorCode == unchecked((int)0x80004005))
                        MessageBox.Show("No se pudo aplicar: parámetro de solo lectura o UCS bloqueado.\n" + ex.Message);
                    else
                        MessageBox.Show("Error COM: " + ex.Message);
                }
                catch (Exception ex)
                {
                    if (tx != null) tx.Abort();
                    MessageBox.Show("Error al aplicar: " + ex.Message);
                }
            });
        }



        private void ApplyOne(Inventor.Parameter p, string expr, bool isAngle)
        {
            if (p == null) return;
            expr = (expr ?? "").Trim();
            if (string.IsNullOrWhiteSpace(expr)) return;

            try
            {
                // Si el usuario escribe una expresión "cruda", déjala tal cual
                if (LooksLikeRawExpression(expr))
                {
                    if (expr.StartsWith("=")) expr = expr.Substring(1).Trim();
                    p.Expression = expr;      // Inventor la evalúa
                    return;
                }

                // Caso relativo/absoluto con números
                var units = UnitEnum(GetUnits(p), isAngle);
                double cur = GetVal(p);

                char op = expr[0];
                if (op == '+' || op == '-' || op == '*' || op == '/')
                {
                    string rhs = expr.Substring(1).Trim();
                    double result;

                    if (op == '+' || op == '-')
                    {
                        // Delta con unidades: "+10", "-2.5 mm", etc.
                        double delta = Convert.ToDouble(_uom.GetValueFromExpression(rhs, units));
                        result = (op == '+') ? cur + delta : cur - delta;
                    }
                    else
                    {
                        // Factor puro: "*1.1" o "/2"
                        double factor = double.Parse(rhs, CultureInfo.InvariantCulture);
                        result = (op == '*') ? cur * factor : cur / factor;
                    }

                    // SIEMPRE por Expression con unidades para evitar E_INVALIDARG
                    string txt = _uom.GetStringFromValue(result, units); // "123 mm" / "10 deg"
                    p.Expression = txt;
                }
                else
                {
                    // Absoluto con unidades (o sin ellas pero con UnitsType)
                    double abs = Convert.ToDouble(_uom.GetValueFromExpression(expr, units));
                    string txt = _uom.GetStringFromValue(abs, units);
                    p.Expression = txt;
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("Inventor rechazó el valor para '" + SafeName(p) + "'.\n" + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Valor inválido: " + expr + "\n\n" + ex.Message);
            }
        }


        // --- NUDGES (pasos) + HOTKEYS + UNDO ---

        // Nudge unificado: delta en mm (offset) o deg (ángulos)
        private void Nudge(Inventor.Parameter p, double delta, bool isAngle)
        {
            if (p == null || ActiveDoc == null) return;

            void Body()
            {
                var tm = _app?.TransactionManager;
                Inventor.Transaction tx = null;

                try
                {
                    if (tm != null)
                        tx = tm.StartTransaction((Inventor._Document)ActiveDoc, "UCS nudge");

                    // 1) Lee valor ACTUAL en unidades internas (cm / rad)
                    double curInternal = GetVal(p);

                    // 2) Convierte el delta de UI -> internas
                    //    mm -> cm  ( /10 )    |   deg -> rad (*PI/180)
                    double incInternal = isAngle
                        ? (delta * Math.PI / 180.0)
                        : (delta / 10.0);

                    double nextInternal = curInternal + incInternal;

                    // 3) Aplica en internas (SetVal ya hace fallback a Expression si hiciera falta)
                    SetVal(p, nextInternal, isAngle);

                    try { ActiveDoc.Update(); } catch { /* ignora si no se puede */ }
                    tx?.End();
                }
                catch
                {
                    try { tx?.Abort(); } catch { }
                    throw;
                }
            }

            // Usa tu wrapper de “modo rápido” si existe
            try { WithFast(Body); } catch { Body(); }

            // Refresca la UI
            LoadValues();
        }




        // UcsForm.cs
        // --- ÚNICO override ---
        // --- ÚNICO override ---
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (HandleHotkey(keyData)) return true;
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // Helper: ¿estamos editando texto/números?
        private bool IsEditingTextual()
        {
            var c = this.ActiveControl;
            if (c == null) return false;
            if (c is TextBoxBase) return true;
            if (c is NumericUpDown) return true;
            return false;
        }

        // --- ÚNICO router de hotkeys ---
        private bool HandleHotkey(Keys keyData)
        {
            bool editing = IsEditingTextual();

            // 0) Recientes (funcionan aunque _ucs sea null)
            if (_cfg != null && _cfg.RecentHotkeys != null)
            {
                int max = Math.Min(6, _cfg.RecentHotkeys.Length);
                for (int i = 0; i < max; i++)
                {
                    string g = _cfg.RecentHotkeys[i];
                    if (!string.IsNullOrEmpty(g) && HotkeyUtil.Matches(keyData, g))
                        return SelectRecentByIndex(i);
                }
            }

            // 1) Pick (configurable)
            if (_cfg != null && !string.IsNullOrEmpty(_cfg.PickHotkey) &&
                HotkeyUtil.Matches(keyData, _cfg.PickHotkey))
            { TryPickUcs(); return true; }

            // 2) Atajos UX generales
            // Ctrl+Enter = Siguiente UCS (no rompe Enter normal)
            if (keyData == (Keys.Control | Keys.Enter)) { NextUcs(); return true; }

            // Alt+Enter = Aplicar (mismo efecto que click en btnApply)
            if (keyData == (Keys.Alt | Keys.Enter))
            {
                if (btnApply != null && btnApply.Enabled) btnApply.PerformClick();
                return true;
            }

            // Enter simple: que lo maneje AcceptButton (no interceptar mientras editas)
            if (keyData == Keys.Enter && editing) return false;

            // 3) Copiar / Pegar (no secuestrar si estás editando)
            if (keyData == (Keys.Control | Keys.C))
            {
                if (!editing) { if (btnCopy != null) btnCopy.PerformClick(); return true; }
                return false;
            }
            if (keyData == (Keys.Control | Keys.V))
            {
                if (!editing) { if (btnPaste != null) btnPaste.PerformClick(); return true; }
                return false;
            }

            // 4) Medir → llenar nudge (siempre disponibles)
            if (keyData == (Keys.Alt | Keys.D)) { MeasureDistanceToStepMm(); return true; } // Distancia → mm
            if (keyData == (Keys.Alt | Keys.G)) { MeasureAngleToStepDeg(); return true; } // Ángulo → grados

            // 5) Desde aquí sí necesitamos un _ucs y tags válidos
            if (_ucs == null) return false;

            double stepMM = (numStepMm != null) ? (double)numStepMm.Value : 1.0;
            double stepDeg = (numStepDeg != null) ? (double)numStepDeg.Value : 1.0;

            var px = (txtXo != null) ? txtXo.Tag as Inventor.Parameter : null;
            var py = (txtYo != null) ? txtYo.Tag as Inventor.Parameter : null;
            var pz = (txtZo != null) ? txtZo.Tag as Inventor.Parameter : null;
            var ax = (txtXa != null) ? txtXa.Tag as Inventor.Parameter : null;
            var ay = (txtYa != null) ? txtYa.Tag as Inventor.Parameter : null;
            var az = (txtZa != null) ? txtZa.Tag as Inventor.Parameter : null;

            // 6) Nudges de traslación (Alt+Q/W/A/S/Z/X)
            switch (keyData)
            {
                case (Keys.Alt | Keys.Q): if (px != null) { Nudge(px, -stepMM, false); return true; } break;
                case (Keys.Alt | Keys.W): if (px != null) { Nudge(px, stepMM, false); return true; } break;
                case (Keys.Alt | Keys.A): if (py != null) { Nudge(py, -stepMM, false); return true; } break;
                case (Keys.Alt | Keys.S): if (py != null) { Nudge(py, stepMM, false); return true; } break;
                case (Keys.Alt | Keys.Z): if (pz != null) { Nudge(pz, -stepMM, false); return true; } break;
                case (Keys.Alt | Keys.X): if (pz != null) { Nudge(pz, stepMM, false); return true; } break;
            }

            // 7) Nudges de rotación (Ctrl+Alt+1..6) + NumPad
            switch (keyData)
            {
                // Teclas superiores
                case (Keys.Control | Keys.Alt | Keys.D1): if (ax != null) { Nudge(ax, stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.D2): if (ax != null) { Nudge(ax, -stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.D3): if (ay != null) { Nudge(ay, stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.D4): if (ay != null) { Nudge(ay, -stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.D5): if (az != null) { Nudge(az, stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.D6): if (az != null) { Nudge(az, -stepDeg, true); return true; } break;

                // Numpad
                case (Keys.Control | Keys.Alt | Keys.NumPad1): if (ax != null) { Nudge(ax, stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.NumPad2): if (ax != null) { Nudge(ax, -stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.NumPad3): if (ay != null) { Nudge(ay, stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.NumPad4): if (ay != null) { Nudge(ay, -stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.NumPad5): if (az != null) { Nudge(az, stepDeg, true); return true; } break;
                case (Keys.Control | Keys.Alt | Keys.NumPad6): if (az != null) { Nudge(az, -stepDeg, true); return true; } break;

                // Vista alineada al UCS
                case (Keys.Alt | Keys.V):
                    AlignViewToUcs(_ucs);
                    return true;
            }

            return false;
        }




        private void ShowHelp()
        {
            string logPath = SysPath.Combine(
                SysEnv.GetFolderPath(SysEnv.SpecialFolder.LocalApplicationData),
                "UcsInspectorperu", "addin.log");

            string txt =
        @"UCS Inspector – Ayuda rápida

Selección
• Filtra por nombre o usa 'Pick en pantalla'.
• ENTER = pasar al siguiente UCS de la lista.

Edición por texto (unidades automáticas)
• +10 mm  | -2.5 mm  | *1.02  | /2  | 25 mm
• +5 deg  | -10 deg  | *0.5   | /2  | 30 deg

Nudges (1 transacción por paso)
• Alt+Q / Alt+W → X- / X+
• Alt+A / Alt+S → Y- / Y+
• Alt+Z / Alt+X → Z- / Z+
• Alt+1 / Alt+2 → RX+ / RX-
• Alt+3 / Alt+4 → RY+ / RY-
• Alt+5 / Alt+6 → RZ+ / RZ-

Botones
• Centrar  • Copiar  • Pegar  • Siguiente  • Actualizar  • Aplicar  • Ayuda (F1)

Atajo global sugerido
• Asigna 'Ctrl+Shift+U' en: Tools ▶ Customize ▶ Keyboard ▶ 'UCS Inspector'.

Rutas
• Log: " + logPath + @"
• Settings: %LOCALAPPDATA%\UcsInspectorperu\settings.json";

            var dlg = new Form
            {
                Text = "UCS Inspector – Ayuda",
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                Width = 620,
                Height = 520
            };

            var tb = new System.Windows.Forms.TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Consolas", 9.0f),
                Text = txt
            };
            dlg.Controls.Add(tb);

            var pnl = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            var btnOpenLog = new Button { Left = 10, Top = 8, Width = 120, Text = "Abrir log" };
            btnOpenLog.Click += (s, e) =>
            {
                try
                {
                    if (System.IO.File.Exists(logPath))
                        System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + logPath + "\"");
                    else
                        MessageBox.Show("Aún no existe el log:\n" + logPath);
                }
                catch { }
            };
            var btnClose = new Button { Left = 480, Top = 8, Width = 120, Text = "Cerrar" };
            btnClose.Click += (s, e) => dlg.Close();
            pnl.Controls.Add(btnOpenLog);
            pnl.Controls.Add(btnClose);
            dlg.Controls.Add(pnl);

            dlg.ShowDialog(this);
        }


        private void NextUcs()
        {
            if (cboUcs.Items.Count == 0) return;
            int i = Math.Max(0, cboUcs.SelectedIndex);
            cboUcs.SelectedIndex = (i + 1) % cboUcs.Items.Count;
        }

        // --- Copiar / Pegar / Centrar ---

        private (double xo, double yo, double zo, double xa, double ya, double za) CopyUcs()
        {
            double V(Inventor.Parameter p) { return p == null ? 0.0 : GetVal(p); }
            var r = (V((Inventor.Parameter)txtXo.Tag), V((Inventor.Parameter)txtYo.Tag), V((Inventor.Parameter)txtZo.Tag),
                     V((Inventor.Parameter)txtXa.Tag), V((Inventor.Parameter)txtYa.Tag), V((Inventor.Parameter)txtZa.Tag));

            // también al portapapeles del sistema (para pegar en otro documento/ventana)
            try
            {
                string S(Inventor.Parameter p, double val, bool ang)
                {
                    var ut = UnitEnum(GetUnits(p), ang);
                    return _uom.GetStringFromValue(val, ut);
                }
                var pack = new ClipPacket
                {
                    xo = S((Inventor.Parameter)txtXo.Tag, r.Item1, false),
                    yo = S((Inventor.Parameter)txtYo.Tag, r.Item2, false),
                    zo = S((Inventor.Parameter)txtZo.Tag, r.Item3, false),
                    xa = S((Inventor.Parameter)txtXa.Tag, r.Item4, true),
                    ya = S((Inventor.Parameter)txtYa.Tag, r.Item5, true),
                    za = S((Inventor.Parameter)txtZa.Tag, r.Item6, true),
                };
                var json = ToJson(pack);
                if (!string.IsNullOrEmpty(json))
                    Clipboard.SetText(json);

            }
            catch { /* no bloquear */ }

            return r;
        }

        // dentro de TryGetFromClipboard: igual que ahora para convertir a números (pack.xo ...)
        // además devuélvelo aparte si hay expresiones:
        private bool TryGetFromClipboard(
      out (double xo, double yo, double zo, double xa, double ya, double za) v,
      out ClipPacket raw)
        {
            v = (0, 0, 0, 0, 0, 0);
            raw = null;

            try
            {
                if (!Clipboard.ContainsText()) return false;
                var s = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(s)) return false;

                // ¿Tiene forma de nuestro JSON?
                if (s.StartsWith("{") && s.IndexOf("\"kind\"", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ClipPacket pkt;
                    if (TryFromJson<ClipPacket>(s, out pkt) &&
                        pkt != null && !string.IsNullOrEmpty(pkt.kind) &&
                        pkt.kind.StartsWith("UCS-", StringComparison.OrdinalIgnoreCase))
                    {
                        raw = pkt;

                        // parsea "92 mm"/"10 deg" → dobles (unidades internas)
                        double P(string text, bool ang)
                        {
                            if (string.IsNullOrWhiteSpace(text)) return 0.0;
                            var ut = ang ? Inventor.UnitsTypeEnum.kDegreeAngleUnits
                                         : Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
                            return System.Convert.ToDouble(_uom.GetValueFromExpression(text, ut));
                        }

                        v = (
                            P(pkt.xo, false),
                            P(pkt.yo, false),
                            P(pkt.zo, false),
                            P(pkt.xa, true),
                            P(pkt.ya, true),
                            P(pkt.za, true)
                        );
                        return true;
                    }
                }

                // no es nuestro formato → no lo procesamos
                return false;
            }
            catch { return false; }
        }



        private void PasteUcs((double xo, double yo, double zo, double xa, double ya, double za) v)
        {
            void Set(Inventor.Parameter p, double val, bool ang)
            {
                if (p == null) return;
                SetVal(p, val, ang);
            }
            Set((Inventor.Parameter)txtXo.Tag, v.xo, false);
            Set((Inventor.Parameter)txtYo.Tag, v.yo, false);
            Set((Inventor.Parameter)txtZo.Tag, v.zo, false);
            Set((Inventor.Parameter)txtXa.Tag, v.xa, true);
            Set((Inventor.Parameter)txtYa.Tag, v.ya, true);
            Set((Inventor.Parameter)txtZa.Tag, v.za, true);
            ActiveDoc.Update();
            LoadValues();
        }

        // Clase auxiliar (puede ir en el mismo .cs)
        internal class NoFlickerPanel : System.Windows.Forms.Panel
        {
            public NoFlickerPanel()
            {
                this.DoubleBuffered = true;
                this.ResizeRedraw = true;
            }
        }

        private void CenterOnUcs(Inventor.UserCoordinateSystem u)
        {
            if (u == null) return;
            try
            {
                var ss = ActiveDoc.SelectSet;
                ss.Clear();
                ss.Select(u);
                _app.ActiveView.Fit(true);
            }
            catch { }
        }

        // -------- Utilidades COM --------

        private Inventor.Parameter TryGetParam(object comObj, string propName)
        {
            try
            {
                var val = comObj.GetType().InvokeMember(
                  propName, BindingFlags.GetProperty, null, comObj, new object[0]);
                return val as Inventor.Parameter;
            }
            catch { return null; }
        }

        private Inventor.Parameter FindParamByName(Inventor.UserCoordinateSystem ucs, string axis, bool isAngle)
        {
            Inventor.Parameters parameters = GetParametersFromDoc(ActiveDoc);
            if (parameters == null) return null;

            var keys = isAngle
              ? new[] { axis + "Angle", "Angle " + axis, "Ángulo " + axis, "Rotación " + axis, "Rot " + axis }
              : new[] { axis + "Offset", "Offset " + axis, "Desplazamiento " + axis, "Delta " + axis, axis + " Offset" };

            var list = ToList(parameters);

            var withUcs = list.Where(p =>
            {
                var n = SafeName(p);
                return n.IndexOf(ucs.Name, StringComparison.OrdinalIgnoreCase) >= 0 &&
                       keys.Any(k => n.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);
            }).ToList();

            var candidates = withUcs.Count > 0 ? withUcs :
                             list.Where(p => keys.Any(k => SafeName(p).IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();

            foreach (var p in candidates)
            {
                string u = (GetUnits(p) ?? "").ToLowerInvariant();
                if (isAngle && (u.Contains("deg") || u.Contains("rad"))) return p;
                if (!isAngle && (u.Contains("mm") || u.Contains("cm") || u.Contains("length"))) return p;
            }
            return candidates.Count > 0 ? candidates[0] : null;
        }

        private Inventor.Parameters GetParametersFromDoc(Inventor.Document doc)
        {
            if (doc is Inventor.PartDocument) return ((Inventor.PartDocument)doc).ComponentDefinition.Parameters;
            if (doc is Inventor.AssemblyDocument) return ((Inventor.AssemblyDocument)doc).ComponentDefinition.Parameters;
            return null;
        }

        private List<Inventor.Parameter> ToList(Inventor.Parameters ps)
        {
            var list = new List<Inventor.Parameter>();
            if (ps == null) return list;
            for (int i = 1; i <= ps.Count; i++) list.Add(ps[i]);
            return list;
        }

        private string SafeName(Inventor.Parameter p)
        {
            try { return p.Name ?? ""; } catch { return ""; }
        }

        private string GetUnits(Inventor.Parameter p)
        {
            try
            {
                object u = p.get_Units();          // <- accessor COM
                return u as string ?? (u != null ? u.ToString() : "");
            }
            catch
            {
                // Fallback por si en otra máquina sí hay propiedad Units
                try
                {
                    var prop = p.GetType().GetProperty("Units");
                    if (prop != null)
                    {
                        var v = prop.GetValue(p, null);
                        return v != null ? v.ToString() : "";
                    }
                }
                catch { }
                return "";
            }
        }

        private static bool LooksLikeRawExpression(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            if (s.StartsWith("=")) return true;
            // ¿hay letras o paréntesis? (d0, mm, deg, cos(), etc.)
            foreach (char ch in s)
                if (char.IsLetter(ch) || ch == '(' || ch == ')') return true;
            return false;
        }

        // Fallback seguro: primero Value, si E_FAIL → Expression con unidades


        private void SetVal(Inventor.Parameter p, double internalValue, bool isAngle)
        {
            if (p == null) return;

            try
            {
                // Camino preferido: asignar valor interno (cm o rad)
                p.Value = internalValue;
                return;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Si el parámetro está “con expresión”, escribir como expresión con unidades de usuario
                try
                {
                    // Convierte internalValue -> texto con unidades de usuario (mm/deg)
                    var ut = TargetInternalUnits(isAngle);
                    // _uom.GetStringFromValue espera valor interno + unidad destino
                    string exprText = _uom != null
                        ? _uom.GetStringFromValue(internalValue, ut)           // p.ej. "50 mm" / "10 deg"
                        : (internalValue.ToString(CultureInfo.InvariantCulture) + " " + DefaultUserUnitSuffix(isAngle));

                    p.Expression = exprText;
                }
                catch { /* ignora */ }
            }
        }





        private static double GetVal(Inventor.Parameter p)
        {
            if (p == null) return 0.0;
            try { return Convert.ToDouble(p.Value); }
            catch { return 0.0; }
        }

        private void SetVal(Inventor.Parameter p, double v)
        {
            try { p.Value = v; } catch { /* read-only → se informará arriba */ }
        }

        private Inventor.UnitsTypeEnum UnitEnum(string units, bool isAngleDefault)
        {
            string u = (units ?? "").Trim().ToLowerInvariant();
            switch (u)
            {
                case "mm": return Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
                case "cm": return Inventor.UnitsTypeEnum.kCentimeterLengthUnits;
                case "in":
                case "inch":
                case "inches": return Inventor.UnitsTypeEnum.kInchLengthUnits;
                case "deg":
                case "degree":
                case "degrees": return Inventor.UnitsTypeEnum.kDegreeAngleUnits;
                case "rad":
                case "radian":
                case "radians": return Inventor.UnitsTypeEnum.kRadianAngleUnits;
                case "ul":
                case "unitless": return Inventor.UnitsTypeEnum.kUnitlessUnits;
                default:
                    return isAngleDefault
                        ? Inventor.UnitsTypeEnum.kDegreeAngleUnits
                        : Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
            }




        }

        private void CopyUcsExpressionsToClipboard()
        {
            var pack = new ClipPacket { kind = "UCS-expr", ver = 2 };

            string Val(Inventor.Parameter p, bool ang)
            {
                if (p == null) return null;
                var ut = UnitEnum(GetUnits(p), ang);
                return _uom.GetStringFromValue(GetVal(p), ut); // "92 mm", "10 deg"
            }
            string Exp(Inventor.Parameter p)
            {
                if (p == null) return null;
                try { return p.Expression; } catch { return null; }
            }

            // valores con unidades (compatibilidad)
            pack.xo = Val((Inventor.Parameter)txtXo.Tag, false);
            pack.yo = Val((Inventor.Parameter)txtYo.Tag, false);
            pack.zo = Val((Inventor.Parameter)txtZo.Tag, false);
            pack.xa = Val((Inventor.Parameter)txtXa.Tag, true);
            pack.ya = Val((Inventor.Parameter)txtYa.Tag, true);
            pack.za = Val((Inventor.Parameter)txtZa.Tag, true);

            // expresiones (si existen)
            pack.xoExp = Exp((Inventor.Parameter)txtXo.Tag);
            pack.yoExp = Exp((Inventor.Parameter)txtYo.Tag);
            pack.zoExp = Exp((Inventor.Parameter)txtZo.Tag);
            pack.xaExp = Exp((Inventor.Parameter)txtXa.Tag);
            pack.yaExp = Exp((Inventor.Parameter)txtYa.Tag);
            pack.zaExp = Exp((Inventor.Parameter)txtZa.Tag);

            var json = ToJson(pack);
            if (!string.IsNullOrEmpty(json))
                Clipboard.SetText(json);
        }



        private const string PlaceholderText = "ej. +10 mm / *1.1 / 25 mm";

        private static bool IsPlaceholderText(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return true;
            s = s.Trim();
            return s.IndexOf("ej.", StringComparison.OrdinalIgnoreCase) >= 0 || s == PlaceholderText;
        }



        private bool ApplyOneFromBox(System.Windows.Forms.TextBox tb, bool isAngle)
        {
            var p = tb?.Tag as Inventor.Parameter;
            if (p == null) return false;

            string raw = (tb.Text ?? "").Trim();
            if (raw.Length == 0) return false;

            // 1) Valor actual en UNIDADES DE USUARIO (mm/deg) para que TryEvalEdit funcione como siempre
            double curUser = isAngle ? (GetVal(p) * 180.0 / Math.PI)   // rad -> deg
                                     : (GetVal(p) * 10.0);             // cm  -> mm

            double targetInternal;

            // 2) Intenta evaluar la edición relativa/absoluta (en mm/deg)
            if (TryEvalEdit(raw, curUser, isAngle, out double targetUser))
            {
                // mm/deg -> interno
                targetInternal = isAngle ? (targetUser * Math.PI / 180.0) : (targetUser / 10.0);
            }
            else
            {
                // 3) Fallbacks: "= 25 mm" | "=10 deg" | "12.34" | "25 mm" | "10deg", etc.
                string s = raw.StartsWith("=") ? raw.Substring(1).Trim() : raw;

                // 3a) número simple (en mm/deg)
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out targetUser) ||
                    double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out targetUser))
                {
                    targetInternal = isAngle ? (targetUser * Math.PI / 180.0) : (targetUser / 10.0);
                }
                else
                {
                    // 3b) con UoM (acepta "25 mm", "10deg", etc.) -> devuelve unidades internas (cm/rad)
                    if (_uom == null) return false;

                    try
                    {
                        bool hasUnits = s.Any(char.IsLetter);
                        string expr = hasUnits ? s : (s + (isAngle ? " deg" : " mm"));
                        targetInternal = Convert.ToDouble(_uom.GetValueFromExpression(expr, Type.Missing));
                    }
                    catch { return false; }
                }
            }

            // 4) Aplica en INTERNAS (cm/rad). SetVal se encarga del fallback por Expression si hace falta.
            try
            {
                SetVal(p, targetInternal, isAngle);
                try { ActiveDoc?.Update(); } catch { }
                LoadValues();
                tb.SelectAll();
                return true;
            }
            catch { return false; }
        }





        private void Log(string msg)
        {
            try
            {
                string dir = SysPath.Combine(SysEnv.GetFolderPath(SysEnv.SpecialFolder.LocalApplicationData), "UcsInspectorperu"); if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                string file = SysPath.Combine(dir, "addin.log");
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + msg + SysEnv.NewLine;
                SysFile.AppendAllText(file, line);
            }
            catch
            {
                // fallback al TEMP si falla LocalAppData
                try
                {
                    string file = SysPath.Combine(SysEnv.GetEnvironmentVariable("TEMP") ?? SysPath.GetTempPath(), "UcsInspectorperu.addin.log");
                    string temp = SysEnv.GetEnvironmentVariable("TEMP") ?? SysEnv.GetFolderPath(SysEnv.SpecialFolder.LocalApplicationData);
                }
                catch { /* silenciar */ }
            }
        }


        private void PasteAsDelta(double xo, double yo, double zo, double xa, double ya, double za)
        {
            WithFast(() =>
            {
                var tm = _app.TransactionManager;
                Inventor.Transaction tx = null;
                try
                {
                    tx = tm.StartTransaction((Inventor._Document)ActiveDoc, "UCS paste delta");

                    void Add(WinTextBox tb, double dv, bool ang)
                    {
                        var p = (Inventor.Parameter)tb.Tag;
                        if (p == null || IsReadOnly(p)) return;

                        try
                        {
                            SetVal(p, GetVal(p) + dv, ang); // suma el delta
                        }
                        catch (System.Runtime.InteropServices.COMException ex)
                        {
                            // E_FAIL o similar → lo saltamos
                            Log("PasteΔ: COMEx en " + p.Name + ": " + ex.Message);
                        }
                    }

                    Add(txtXo, xo, false);
                    Add(txtYo, yo, false);
                    Add(txtZo, zo, false);
                    Add(txtXa, xa, true);
                    Add(txtYa, ya, true);
                    Add(txtZa, za, true);

                    ActiveDoc.Update();
                    tx.End();
                    LoadValues();
                }
                catch
                {
                    if (tx != null) tx.Abort();
                    throw;
                }
            });
        }
        private void btnPasteDelta_Click(object sender, EventArgs e)
        {
            string err;
            double xo, yo, zo, xa, ya, za;
            if (!TryReadClipboardDelta(out xo, out yo, out zo, out xa, out ya, out za, out err))
            {
                MessageBox.Show("Pegar Δ: " + err, "UCS Inspector", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            PasteAsDelta(xo, yo, zo, xa, ya, za);
        }



        private bool TryReadClipboardDelta(out double xo, out double yo, out double zo,
                                   out double xa, out double ya, out double za,
                                   out string err)
        {
            xo = yo = zo = xa = ya = za = 0; err = null;
            try
            {
                if (!Clipboard.ContainsText()) { err = "El portapapeles no tiene texto."; return false; }
                var raw = Clipboard.GetText().Trim();
                if (string.IsNullOrEmpty(raw)) { err = "El portapapeles está vacío."; return false; }

                // normaliza separadores
                var s = raw.Replace('\t', ' ')
                           .Replace(';', ' ')
                           .Replace('\n', ' ')
                           .Replace('\r', ' ');
                while (s.IndexOf("  ", System.StringComparison.Ordinal) >= 0) s = s.Replace("  ", " ");

                var tok = s.Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries);
                if (tok.Length < 6) { err = "Se esperan 6 números (xo yo zo xa ya za)."; return false; }

                double[] v = new double[6];
                for (int i = 0; i < 6; i++)
                {
                    var t = tok[i].ToLowerInvariant();
                    t = t.Replace("mm", "").Replace("deg", "").Trim();
                    // admite coma decimal
                    if (t.IndexOf(',') >= 0 && t.IndexOf('.') < 0) t = t.Replace(',', '.');

                    double d;
                    if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                                         System.Globalization.CultureInfo.InvariantCulture, out d))
                    {
                        err = "No pude interpretar el valor: " + tok[i];
                        return false;
                    }
                    v[i] = d;
                }

                xo = v[0]; yo = v[1]; zo = v[2];
                xa = v[3]; ya = v[4]; za = v[5];
                return true;
            }
            catch (System.Exception ex)
            {
                err = ex.Message; return false;
            }
        }


        private ToolTip EnsureTooltip()
        {
            if (_tt == null)
            {
                _tt = new ToolTip
                {
                    AutoPopDelay = 5000,
                    InitialDelay = 300,
                    ReshowDelay = 100,
                    ShowAlways = true
                };
            }
            return _tt;
        }




     

        private void InitRecentsFromConfig()
        {
            _mru = new System.Collections.Generic.List<string>(_cfg.RecentUcs ?? new string[0]);
            _fav = new System.Collections.Generic.List<string>(_cfg.FavoriteUcs ?? new string[0]);
        }

        // Llamar esto cada vez que "usas" un UCS (selección, Pick, Siguiente, Aplicar…)
        private void TouchRecent(string name)
        {
            if (string.IsNullOrEmpty(name)) return;

            // quita si ya estaba y pon al frente
            _mru.RemoveAll(n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase));
            _mru.Insert(0, name);

            // recorta capacidad
            int cap = (_cfg.RecentCapacity > 0) ? _cfg.RecentCapacity : 24;
            while (_mru.Count > cap) _mru.RemoveAt(_mru.Count - 1);

            // persiste
            _cfg.RecentUcs = _mru.ToArray();
            _cfg.Save();
            // usa tu método existente
        }

        // lista para UI: Favoritos (en su orden) + MRU (excluyendo favoritos), truncado a RecentMax
        private string[] GetRecentDisplayNames()
        {
            var outList = new System.Collections.Generic.List<string>();
            int max = (_cfg.RecentMax > 0) ? _cfg.RecentMax : 6;

            // favoritos primero (sólo si existen en el doc—opcional filtrar luego)
            for (int i = 0; i < _fav.Count && outList.Count < max; i++)
            {
                var n = _fav[i];
                if (!string.IsNullOrEmpty(n) && !outList.Contains(n)) outList.Add(n);
            }

            // luego MRU no favoritos
            for (int i = 0; i < _mru.Count && outList.Count < max; i++)
            {
                var n = _mru[i];
                if (string.IsNullOrEmpty(n)) continue;
                bool isFav = _fav.Exists(f => string.Equals(f, n, StringComparison.OrdinalIgnoreCase));
                if (!isFav && !outList.Contains(n)) outList.Add(n);
            }

            return outList.ToArray();
        }

        private bool IsFavorite(string name)
        {
            for (int i = 0; i < _fav.Count; i++)
                if (string.Equals(_fav[i], name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        private void SetFavorite(string name, bool makeFav)
        {
            if (string.IsNullOrEmpty(name)) return;

            // quita duplicados
            _fav.RemoveAll(n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase));

            if (makeFav) _fav.Add(name); // se fija al final (más a la derecha)
            _cfg.FavoriteUcs = _fav.ToArray();
            _cfg.Save();

            RefreshRecentButtons();
        }



        private void ApplySettingsToUi(UiSettings s)
        {
            try
            {
                // tamaño/posición
                Left = s.Left; Top = s.Top;
                Width = s.Width; Height = s.Height;

                // pasos
                if (numStepMm != null) numStepMm.Value = (decimal)s.StepMm;
                if (numStepDeg != null) numStepDeg.Value = (decimal)s.StepDeg;

                // pegar Δ
                if (chkDelta != null) chkDelta.Checked = s.PasteAsDelta;

                // garantiza mínimos
                if (s.RecentMax < 6) s.RecentMax = 6;
            }
            catch { }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try { SaveSettingsFromUi(); } catch { }
            try { GlobalHotkeys.UnregisterAll(); } catch { }
            base.OnFormClosed(e);
        }



        // UcsForm.cs
        private void ReloadSettingsAndHotkeys()
        {
            try
            {
                if (_cfg == null)
                    _cfg = UiSettings.Load() ?? new UiSettings();

                // ── UI (no repintar si estás editando un TextBox/NumericUpDown) ─────────
                bool editing = IsEditingTextual(); // helper: ActiveControl es TextBoxBase/NumericUpDown
                if (!editing)
                {
                    ApplySettingsToUi(_cfg);
                    InitRecentsFromConfig();
                    RefreshRecentButtons();
                }

                // ── Hotkeys globales ────────────────────────────────────────────────────
                GlobalHotkeys.UnregisterAll();

                var list = new List<Tuple<uint, uint, Action>>();

                // 0) Toggle (Alt+F12 por defecto) → SIEMPRE disponible
                HotkeyUtil.Parsed tg;
                string toggle = (!string.IsNullOrEmpty(_cfg.ToggleHotkey)) ? _cfg.ToggleHotkey : "Alt+F12";
                if (HotkeyUtil.TryParse(toggle, out tg) && tg.IsValid)
                    list.Add(Tuple.Create(tg.FsModifiers, tg.VirtualKey, (Action)SafeToggleShow));

                // 1) Resto de globales solo si está habilitado en settings
                if (_cfg.GlobalNudgeHotkeys)
                {
                    // Recientes (1..6)
                    var rh = _cfg.RecentHotkeys;
                    int max = (rh != null) ? Math.Min(6, rh.Length) : 0;
                    for (int i = 0; i < max; i++)
                    {
                        HotkeyUtil.Parsed p;
                        if (!string.IsNullOrEmpty(rh[i]) &&
                            HotkeyUtil.TryParse(rh[i], out p) && p.IsValid)
                        {
                            int captured = i;
                            list.Add(Tuple.Create(p.FsModifiers, p.VirtualKey,
                                (Action)(() => SelectRecentByIndex(captured))));
                        }
                    }

                    // Pick en pantalla
                    HotkeyUtil.Parsed pp;
                    if (!string.IsNullOrEmpty(_cfg.PickHotkey) &&
                        HotkeyUtil.TryParse(_cfg.PickHotkey, out pp) && pp.IsValid)
                    {
                        list.Add(Tuple.Create(pp.FsModifiers, pp.VirtualKey, (Action)TryPickUcs));
                    }
                }

                // Registrar si hay algo y la ventana de Inventor existe
                if (list.Count > 0 && _app != null && _app.MainFrameHWND != 0)
                    GlobalHotkeys.RegisterMany(new IntPtr(_app.MainFrameHWND), list);
            }
            catch (Exception ex)
            {
                // TODO: SafeLog("ReloadSettingsAndHotkeys: " + ex.Message);
            }
        }




        private void WireHotkeysToChildren(Control root)
        {
            if (root == null) return;
            root.KeyDown += (s, e) =>
            {
                if (HandleHotkey(e.KeyData)) { e.Handled = true; e.SuppressKeyPress = true; }
            };
            foreach (Control c in root.Controls) WireHotkeysToChildren(c);
        }

        // Wrapper único: toda la lógica de atajos aquí

       

        // Mantén ProcessCmdKey como puerta de entrada principal
        

        private void BtnPaste_Click(object sender, EventArgs e)
        {
            (double xo, double yo, double zo, double xa, double ya, double za) vals;
            ClipPacket clip;

            if (TryGetFromClipboard(out vals, out clip))
            {
                if (chkDelta != null && chkDelta.Checked)
                {
                    // Δ: suma al valor actual
                    PasteAsDelta(vals.xo, vals.yo, vals.zo, vals.xa, vals.ya, vals.za);
                }
                else
                {
                    // Absoluto (si tienes expresiones en clip, tu PasteUcs ya las maneja/ignora)
                    PasteUcs(vals);
                }
                LoadValues();
                return;
            }

            if (_clipboard.HasValue)
            {
                if (chkDelta != null && chkDelta.Checked)
                    PasteAsDelta(_clipboard.Value.xo, _clipboard.Value.yo, _clipboard.Value.zo,
                                 _clipboard.Value.xa, _clipboard.Value.ya, _clipboard.Value.za);
                else
                    PasteUcs(_clipboard.Value);

                LoadValues();
                return;
            }

            MessageBox.Show("No hay datos de UCS en el portapapeles.");
        }







        private void BuildLayout()
        {
            Controls.Clear();

            // ===== Contenedor con scroll =====
            _content = new WinPanel { Dock = DockStyle.Fill, AutoScroll = true };
            Controls.Add(_content);

            // ===== Cabecera =====
            _header = new WinTable
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 3,
                Padding = new Padding(12, 12, 12, 0)
            };
            _header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45f));   // filtro
            _header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45f));   // combo
            _header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140f)); // pick

            txtFilter = new WinTextBox { Dock = DockStyle.Fill, Margin = new Padding(0) };
            cboUcs = new WinComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new Padding(6, 0, 6, 0) };
            btnPick = new WinButton { Text = "Pick en pantalla", Dock = DockStyle.Fill, Height = 24, Margin = new Padding(0) };
            btnPick.Click += (s, e) => TryPickUcs();

            _header.Controls.Add(txtFilter, 0, 0);
            _header.Controls.Add(cboUcs, 1, 0);
            _header.Controls.Add(btnPick, 2, 0);
            _content.Controls.Add(_header);

            // ===== Recientes =====
            _recents = new WinFlow
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(12, 6, 12, 0),
                WrapContents = false
            };
            _content.Controls.Add(_recents);

            // ===== Grid principal =====
            _grid = new WinTable
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 4,
                Padding = new Padding(12, 8, 12, 0)
            };
            // 0: Label (fijo) | 1: Valor (50%) | 2: Ejemplo (50%) | 3: Nudges (fijo)
            _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110f));
            _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120f));
            _content.Controls.Add(_grid);

            // === Filas offsets ===
            AddRow(_grid, "X Offset", out lblXo, out txtXo, out btnXm, out btnXp, "-X", "+X");
            AddRow(_grid, "Y Offset", out lblYo, out txtYo, out btnYm, out btnYp, "-Y", "+Y");
            AddRow(_grid, "Z Offset", out lblZo, out txtZo, out btnZm, out btnZp, "-Z", "+Z");

            // === Filas ángulos ===
            AddRow(_grid, "X Angle", out lblXa, out txtXa, out btnRxm, out btnRxp, "-RX", "+RX");
            AddRow(_grid, "Y Angle", out lblYa, out txtYa, out btnRym, out btnRyp, "-RY", "+RY");
            AddRow(_grid, "Z Angle", out lblZa, out txtZa, out btnRzm, out btnRzp, "-RZ", "+RZ");

            // === Ejemplos (labels de solo lectura) ===
            _hintX = MakeHint("ej. +10 mm / *1.1 / 25 mm");
            _hintY = MakeHint("ej. +10 mm / *1.1 / 25 mm");
            _hintZ = MakeHint("ej. +10 mm / *1.1 / 25 mm");
            _hintRX = MakeHint("ej. +10 deg / *1.1 / 25 deg");
            _hintRY = MakeHint("ej. +10 deg / *1.1 / 25 deg");
            _hintRZ = MakeHint("ej. +10 deg / *1.1 / 25 deg");

            PutHintOnRow(_hintX, 0);
            PutHintOnRow(_hintY, 1);
            PutHintOnRow(_hintZ, 2);
            PutHintOnRow(_hintRX, 3);
            PutHintOnRow(_hintRY, 4);
            PutHintOnRow(_hintRZ, 5);

            // ===== Acciones =====
            var actions = new WinFlow
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(12, 8, 12, 0)
            };

            // Steps cómodos
            numStepMm = new NumericUpDown { Width = 70, DecimalPlaces = 3, Minimum = 0.001M, Maximum = 10000M, Increment = 0.1M, Value = 1M, Margin = new Padding(0, 0, 6, 0), TextAlign = HorizontalAlignment.Right };
            numStepDeg = new NumericUpDown { Width = 70, DecimalPlaces = 2, Minimum = 0.01M, Maximum = 360M, Increment = 0.1M, Value = 1M, Margin = new Padding(0, 0, 12, 0), TextAlign = HorizontalAlignment.Right };

            actions.Controls.Add(new WinLabel { Text = "Paso mm", AutoSize = true, Margin = new Padding(0, 4, 6, 0) });
            actions.Controls.Add(numStepMm);
            actions.Controls.Add(new WinLabel { Text = "Paso °", AutoSize = true, Margin = new Padding(0, 4, 6, 0) });
            actions.Controls.Add(numStepDeg);

            // Botones principales
            btnCenter = NewCmd("Centrar", (s, e) => CenterOnUcs(_ucs));
            btnCopy = NewCmd("Copiar", (s, e) => _clipboard = CopyUcs());
            btnPaste = NewCmd("Pegar", BtnPaste_Click);
            btnNext = NewCmd("Siguiente", (s, e) => NextUcs());
            btnRefresh = NewCmd("Actualiza", (s, e) => LoadValues());
            var btnExp = NewCmd("Copiar e", (s, e) => CopyUcsExpressionsToClipboard());
            btnHelp = NewCmd("Ayuda (F1)", (s, e) => ShowHelp());
            btnApply = NewCmd("Aplicar", (s, e) => ApplyChanges());
            chkDelta = new CheckBox { Text = "Pegar Δ", AutoSize = true, Margin = new Padding(12, 4, 12, 0) };
            var btnViewUcs = NewCmd("Vista =", (s, e) => AlignViewToUcs(_ucs));

            // NUEVOS: medir y settings
            var btnMeasureMm = NewCmd("Medir mm (Alt+D)", (s, e) => MeasureDistanceToStepMm());
            var btnMeasureDeg = NewCmd("Medir ° (Alt+G)", (s, e) => MeasureAngleToStepDeg());
            var btnSettings = NewCmd("Settings…", btnSettings_Click);

            actions.Controls.AddRange(new SysWin.Control[]
            {
        btnCenter, btnCopy, btnPaste, btnNext, btnRefresh, btnExp, btnHelp,
        btnApply, chkDelta, btnViewUcs, btnMeasureMm, btnMeasureDeg, btnSettings
            });

            _content.Controls.Add(actions);

            // al final de BuildLayout()
            WireNudges();
        }



     


        private WinButton NewCmd(string text, EventHandler onClick)
        {
            var b = new WinButton { Text = text, AutoSize = true, Height = 24, Margin = new Padding(6, 0, 0, 0) };
            b.Click += onClick;
            return b;
        }

        private void AddRow(WinTable grid, string label,
            out WinLabel lbl, out WinTextBox txt, out WinButton bMinus, out WinButton bPlus,
            string minusText, string plusText)
        {
            int row = grid.RowCount;
            grid.RowCount++;
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f)); // alto estándar

            lbl = new WinLabel { Text = label, AutoSize = true, Margin = new Padding(0, 6, 6, 0) };
            txt = new WinTextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 6, 2) }; // altura ~22px

            var nudge = new WinFlow { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0) };
            bMinus = new WinButton { Text = minusText, Width = 44, Height = 24, Margin = new Padding(0, 0, 6, 0) };
            bPlus = new WinButton { Text = plusText, Width = 44, Height = 24, Margin = new Padding(0) };
            nudge.Controls.Add(bMinus);
            nudge.Controls.Add(bPlus);

            grid.Controls.Add(lbl, 0, row);
            grid.Controls.Add(txt, 1, row);
            // la col 2 (ejemplo) se llena luego con PutHintOnRow(...)
            grid.Controls.Add(nudge, 3, row);
        }

        // crea un label para la columna de ejemplos
        private WinLabel MakeHint(string text)
        {
            return new WinLabel { Text = text, AutoSize = true, ForeColor = System.Drawing.Color.Gray, Margin = new Padding(0, 6, 0, 0) };
        }

        // coloca el “hint” en la col 2
        private void PutHintOnRow(WinLabel hint, int row)
        {
            _grid.Controls.Add(hint, 2, row);
        }

        // modo compacto: oculta col 2 (ejemplos) y pasa texto a tooltip de los textbox
        private void UpdateLayoutForWidth()
        {
            if (_grid == null) return;
            int w = ClientSize.Width;

            bool hideHints = w < 820;      // umbral para ocultar "ej. +10 mm..."
            bool compact = w < 700;     // umbral para apretar columnas

            // hints visibles/ocultos
            foreach (var h in new[] { _hintX, _hintY, _hintZ, _hintRX, _hintRY, _hintRZ })
                if (h != null) h.Visible = !hideHints;

            _grid.SuspendLayout();
            _grid.ColumnStyles.Clear();

            if (compact)
            {
                // 3 columnas: Label | Valor (100%) | Nudges
                _grid.ColumnCount = 3;
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));  // label
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));   // valor
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96f));   // nudges
                                                                                   // los hints se ocultan (arriba) así no ocupan celda
            }
            else
            {
                // 4 columnas: Label | Valor (55%) | Ejemplo (160) | Nudges (96)
                _grid.ColumnCount = 4;
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110f));
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55f));
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160f));
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96f));
            }

            _grid.ResumeLayout(true);
        }


        // Usa tu alias si lo tienes: using SysWin = System.Windows.Forms;
        private void SetHintsAsTooltips(bool enable)
        {
            // crea el ToolTip si hace falta
            if (_tt == null)
                _tt = new ToolTip { AutoPopDelay = 5000, InitialDelay = 300, ReshowDelay = 100, ShowAlways = true };

            // helper local
            void set(Control c, Label hint)
            {
                if (c == null) return;
                var text = (enable && hint != null) ? hint.Text : null;
                _tt.SetToolTip(c, text);
            }

            // aplica a tus controles (ajusta nombres si cambian)
            set(txtXo, _hintX); set(txtYo, _hintY); set(txtZo, _hintZ);
            set(txtXa, _hintRX); set(txtYa, _hintRY); set(txtZa, _hintRZ);
        }



        // Devuelve la lista final a pintar: Favoritos (todos) + MRU (recortado a cap)



        private void WireNudges()
        {
            if (btnXm != null) btnXm.Click += (s, e) => Nudge(txtXo?.Tag as Inventor.Parameter, -(double)numStepMm.Value, false);
            if (btnXp != null) btnXp.Click += (s, e) => Nudge(txtXo?.Tag as Inventor.Parameter, +(double)numStepMm.Value, false);

            if (btnYm != null) btnYm.Click += (s, e) => Nudge(txtYo?.Tag as Inventor.Parameter, -(double)numStepMm.Value, false);
            if (btnYp != null) btnYp.Click += (s, e) => Nudge(txtYo?.Tag as Inventor.Parameter, +(double)numStepMm.Value, false);

            if (btnZm != null) btnZm.Click += (s, e) => Nudge(txtZo?.Tag as Inventor.Parameter, -(double)numStepMm.Value, false);
            if (btnZp != null) btnZp.Click += (s, e) => Nudge(txtZo?.Tag as Inventor.Parameter, +(double)numStepMm.Value, false);

            if (btnRxm != null) btnRxm.Click += (s, e) => Nudge(txtXa?.Tag as Inventor.Parameter, -(double)numStepDeg.Value, true);
            if (btnRxp != null) btnRxp.Click += (s, e) => Nudge(txtXa?.Tag as Inventor.Parameter, +(double)numStepDeg.Value, true);

            if (btnRym != null) btnRym.Click += (s, e) => Nudge(txtYa?.Tag as Inventor.Parameter, -(double)numStepDeg.Value, true);
            if (btnRyp != null) btnRyp.Click += (s, e) => Nudge(txtYa?.Tag as Inventor.Parameter, +(double)numStepDeg.Value, true);

            if (btnRzm != null) btnRzm.Click += (s, e) => Nudge(txtZa?.Tag as Inventor.Parameter, -(double)numStepDeg.Value, true);
            if (btnRzp != null) btnRzp.Click += (s, e) => Nudge(txtZa?.Tag as Inventor.Parameter, +(double)numStepDeg.Value, true);
        }


        // Acelera acciones: apaga repintado y avisos, ejecuta, y restaura
        private void WithFast(Action action)
        {
            if (_app == null) { action(); return; }
            bool su = true, so = false;
            try { su = _app.ScreenUpdating; _app.ScreenUpdating = false; } catch { }
            try { so = _app.SilentOperation; _app.SilentOperation = true; } catch { }
            try { action(); }
            finally
            {
                try { _app.SilentOperation = so; } catch { }
                try { _app.ScreenUpdating = su; } catch { }
            }
        }









        // ==== Helpers de unidades ====
        private static UnitsTypeEnum TargetInternalUnits(bool isAngle)
            => isAngle ? UnitsTypeEnum.kRadianAngleUnits : UnitsTypeEnum.kCentimeterLengthUnits;

        private static string DefaultUserUnitSuffix(bool isAngle)   // para construir "10 mm" / "5 deg"
            => isAngle ? "deg" : "mm";

        // Lee el valor interno (cm/rad) del parámetro
       



        // Soporta: +10 / -25 / *1.1 / /2 / absoluto (con o sin unidad: mm/deg)
        private bool TryEvalEdit(string s, double cur, bool isAngle, out double result)
        {
            result = cur;
            s = s.Replace(',', '.').Trim();

            string unit = isAngle ? "deg" : "mm";

            double? parse(string txt)
            {
                txt = txt.Trim();
                if (txt.EndsWith(unit, StringComparison.OrdinalIgnoreCase))
                    txt = txt.Substring(0, txt.Length - unit.Length).Trim();

                if (double.TryParse(txt, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                    return v;
                return null;
            }

            if (s.StartsWith("+") || s.StartsWith("-"))
            {
                var v = parse(s.Substring(1));
                if (v == null) return false;
                result = cur + (s[0] == '+' ? +1 : -1) * v.Value;
                return true;
            }
            if (s.StartsWith("*"))
            {
                var v = parse(s.Substring(1));
                if (v == null) return false;
                result = cur * v.Value;
                return true;
            }
            if (s.StartsWith("/"))
            {
                var v = parse(s.Substring(1));
                if (v == null || Math.Abs(v.Value) < 1e-12) return false;
                result = cur / v.Value;
                return true;
            }

            var abs = parse(s);
            if (abs == null) return false;
            result = abs.Value;                // absoluto
            return true;
        }



        // altura base para entradas (ajusta a tu gusto)
        private int RowH => Math.Max(22, Font.Height + 8);

        private WinTextBox NewInput()
        {
            return new WinTextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 2, 6, 2),
                MinimumSize = new Size(60, RowH)
            };
        }



        private void btnSettings_Click(object sender, EventArgs e)
        {
            using (var dlg = new HotkeySettingsForm(_cfg)) // se edita la misma instancia
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    _cfg.Save();                  // guarda JSON
                    ReloadSettingsAndHotkeys();   // aplica sin reiniciar
                }
            }
        }


        internal class KeyCaptureTextBox : SysWin.TextBox
        {
            protected override void OnKeyDown(SysWin.KeyEventArgs e)
            {
                // Limpiar con Backspace/Delete
                if (e.KeyCode == SysWin.Keys.Back || e.KeyCode == SysWin.Keys.Delete)
                {
                    this.Clear();
                    e.Handled = true; e.SuppressKeyPress = true;
                    return;
                }

                // Ignorar cuando es solo un modificador
                if (e.KeyCode == SysWin.Keys.ControlKey ||
                    e.KeyCode == SysWin.Keys.ShiftKey ||
                    e.KeyCode == SysWin.Keys.Menu)
                {
                    e.Handled = true; e.SuppressKeyPress = true;
                    return;
                }

                string s = FormatHotkey(e);
                if (!string.IsNullOrEmpty(s)) this.Text = s;

                e.Handled = true; e.SuppressKeyPress = true;
                base.OnKeyDown(e);
            }

            private static string FormatHotkey(SysWin.KeyEventArgs e)
            {
                string mods = "";
                if (e.Control) mods = Append(mods, "Ctrl");
                if (e.Alt) mods = Append(mods, "Alt");
                if (e.Shift) mods = Append(mods, "Shift");

                string key = MapKey(e.KeyCode);
                if (string.IsNullOrEmpty(key)) return mods;

                return string.IsNullOrEmpty(mods) ? key : (mods + "+" + key);
            }

            private static string Append(string a, string b) { return string.IsNullOrEmpty(a) ? b : a + "+" + b; }

            private static string MapKey(SysWin.Keys k)
            {
                if (k >= SysWin.Keys.A && k <= SysWin.Keys.Z) return k.ToString();         // A..Z
                if (k >= SysWin.Keys.D0 && k <= SysWin.Keys.D9) return "D" + (k - SysWin.Keys.D0); // D0..D9
                if (k >= SysWin.Keys.NumPad0 && k <= SysWin.Keys.NumPad9) return "NumPad" + (k - SysWin.Keys.NumPad0);
                if (k >= SysWin.Keys.F1 && k <= SysWin.Keys.F24) return "F" + (k - SysWin.Keys.F1 + 1);

                if (k == SysWin.Keys.Space) return "Space";
                if (k == SysWin.Keys.Tab) return "Tab";
                if (k == SysWin.Keys.Escape) return "Esc";
                if (k == SysWin.Keys.Enter) return "Enter";
                return k.ToString();
            }
        }

        public class HotkeySettingsForm : SysWin.Form
        {
            private UiSettings _cfg;

            private SysWin.CheckBox chkGlobal;
            private KeyCaptureTextBox txtToggle;
            private KeyCaptureTextBox txtPick;
            private KeyCaptureTextBox[] txtRec = new KeyCaptureTextBox[6];

            private SysWin.Button btnOk, btnCancel, btnDefaults;

            public HotkeySettingsForm(UiSettings cfg)
            {
                if (cfg == null) throw new ArgumentNullException("cfg");
                _cfg = cfg;

                Text = "Configuración de Atajos – UCS Inspector";
                FormBorderStyle = SysWin.FormBorderStyle.FixedDialog;
                StartPosition = SysWin.FormStartPosition.CenterParent;
                MaximizeBox = false; MinimizeBox = false;
                ClientSize = new SysDraw.Size(480, 330);

                BuildUi();
                LoadFromCfg();

                AcceptButton = btnOk;
                CancelButton = btnCancel;
            }

            private void BuildUi()
            {
                int y = 16;

                chkGlobal = new SysWin.CheckBox
                {
                    Text = "Habilitar atajos globales (funcionan sin foco en el formulario)",
                    Location = new SysDraw.Point(16, y),
                    Size = new SysDraw.Size(440, 24)
                };
                Controls.Add(chkGlobal);
                y += 32;

                var lblToggle = new SysWin.Label
                {
                    Text = "Mostrar/Ocultar (toggle):",
                    Location = new SysDraw.Point(16, y + 3),
                    Size = new SysDraw.Size(170, 20)
                };
                Controls.Add(lblToggle);

                txtToggle = new KeyCaptureTextBox
                {
                    Location = new SysDraw.Point(200, y),
                    Size = new SysDraw.Size(120, 24)
                };
                Controls.Add(txtToggle);
                y += 32;

                var lblPick = new SysWin.Label
                {
                    Text = "Pick en pantalla:",
                    Location = new SysDraw.Point(16, y + 3),
                    Size = new SysDraw.Size(170, 20)
                };
                Controls.Add(lblPick);

                txtPick = new KeyCaptureTextBox
                {
                    Location = new SysDraw.Point(200, y),
                    Size = new SysDraw.Size(120, 24)
                };
                Controls.Add(txtPick);
                y += 36;

                var lblRec = new SysWin.Label
                {
                    Text = "Recientes (1–6):",
                    Location = new SysDraw.Point(16, y + 3),
                    Size = new SysDraw.Size(170, 20)
                };
                Controls.Add(lblRec);

                for (int i = 0; i < 6; i++)
                {
                    var t = new KeyCaptureTextBox
                    {
                        Location = new SysDraw.Point(200 + (i % 3) * 90, y + (i / 3) * 32),
                        Size = new SysDraw.Size(80, 24)
                    };
                    Controls.Add(t);
                    txtRec[i] = t;

                    var cap = new SysWin.Label
                    {
                        Text = (i + 1).ToString(),
                        TextAlign = SysDraw.ContentAlignment.MiddleLeft,
                        Location = new SysDraw.Point(200 + (i % 3) * 90 - 18, y + (i / 3) * 32 + 3),
                        Size = new SysDraw.Size(16, 20)
                    };
                    Controls.Add(cap);
                }

                btnDefaults = new SysWin.Button
                {
                    Text = "Restaurar por defecto",
                    Location = new SysDraw.Point(16, ClientSize.Height - 44),
                    Size = new SysDraw.Size(160, 28)
                };
                btnDefaults.Click += (s, e) => LoadDefaults();
                Controls.Add(btnDefaults);

                btnOk = new SysWin.Button
                {
                    Text = "Aceptar",
                    Location = new SysDraw.Point(ClientSize.Width - 180, ClientSize.Height - 44),
                    Size = new SysDraw.Size(80, 28)
                };
                btnOk.Click += OnOk;
                Controls.Add(btnOk);

                btnCancel = new SysWin.Button
                {
                    Text = "Cancelar",
                    Location = new SysDraw.Point(ClientSize.Width - 92, ClientSize.Height - 44),
                    Size = new SysDraw.Size(80, 28),
                    DialogResult = SysWin.DialogResult.Cancel
                };
                Controls.Add(btnCancel);
            }

            private void LoadFromCfg()
            {
                chkGlobal.Checked = _cfg.GlobalNudgeHotkeys;
                txtToggle.Text = _cfg.ToggleHotkey ?? "";
                txtPick.Text = _cfg.PickHotkey ?? "";

                for (int i = 0; i < 6; i++)
                {
                    string v = (_cfg.RecentHotkeys != null && i < _cfg.RecentHotkeys.Length)
                             ? _cfg.RecentHotkeys[i] : "";
                    txtRec[i].Text = v ?? "";
                }
            }

            private void LoadDefaults()
            {
                chkGlobal.Checked = true;
                txtToggle.Text = "Alt+F12";
                txtPick.Text = "Alt+P";
                txtRec[0].Text = "Alt+D1";
                txtRec[1].Text = "Alt+D2";
                txtRec[2].Text = "Alt+D3";
                txtRec[3].Text = "Alt+D4";
                txtRec[4].Text = "Alt+D5";
                txtRec[5].Text = "Alt+D6";
            }

            private void OnOk(object sender, EventArgs e)
            {
                if (!IsValidOrEmpty(txtToggle.Text) ||
                    !IsValidOrEmpty(txtPick.Text) ||
                    !IsValidOrEmpty(txtRec[0].Text) ||
                    !IsValidOrEmpty(txtRec[1].Text) ||
                    !IsValidOrEmpty(txtRec[2].Text) ||
                    !IsValidOrEmpty(txtRec[3].Text) ||
                    !IsValidOrEmpty(txtRec[4].Text) ||
                    !IsValidOrEmpty(txtRec[5].Text))
                {
                    SysWin.MessageBox.Show(this,
                        "Hay atajos con formato no válido. Revisa los campos en rojo.",
                        "Atajos inválidos", SysWin.MessageBoxButtons.OK, SysWin.MessageBoxIcon.Warning);
                    return;
                }

                _cfg.GlobalNudgeHotkeys = chkGlobal.Checked;
                _cfg.ToggleHotkey = (txtToggle.Text ?? "").Trim();
                _cfg.PickHotkey = (txtPick.Text ?? "").Trim();

                _cfg.RecentHotkeys = new string[6];
                for (int i = 0; i < 6; i++)
                    _cfg.RecentHotkeys[i] = (txtRec[i].Text ?? "").Trim();

                DialogResult = SysWin.DialogResult.OK;
                Close();
            }

            private bool IsValidOrEmpty(string s)
            {
                if (string.IsNullOrEmpty(s)) return true;
                HotkeyUtil.Parsed p;
                bool ok = HotkeyUtil.TryParse(s, out p) && p.IsValid;
                if (!ok)
                {
                    // marcar visual al primero que coincida
                    foreach (SysWin.Control c in Controls)
                    {
                        var t = c as SysWin.TextBox;
                        if (t != null && string.Equals(t.Text, s, StringComparison.Ordinal))
                        {
                            t.BackColor = SysDraw.Color.MistyRose;
                            break;
                        }
                    }
                }
                return ok;
            }
        }

     

    }
}
