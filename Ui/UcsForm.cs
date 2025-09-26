using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Inventor;

// Alias para evitar choque con Inventor.Color/TextBox
using DrawColor = System.Drawing.Color;
using WinTextBox = System.Windows.Forms.TextBox;

namespace UcsInspectorperu.Ui
{
    public class UcsForm : Form
    {
        private readonly Inventor.Application _app;
        private ComboBox cboUcs;
        private Label lblXo, lblYo, lblZo, lblXa, lblYa, lblZa;
        private WinTextBox txtXo, txtYo, txtZo, txtXa, txtYa, txtZa;
        private Button btnApply, btnRefresh, btnPick;

        // NUEVO: UI y estado
        private NumericUpDown numStepMm, numStepDeg;
        private Button btnCenter, btnCopy, btnPaste, btnNext;
        private System.Windows.Forms.TextBox txtFilter;
        private readonly List<Inventor.UserCoordinateSystem> _allUcs = new List<Inventor.UserCoordinateSystem>();
        private (double xo, double yo, double zo, double xa, double ya, double za)? _clipboard;

        // NUEVO: nudge buttons
        private Button btnXm, btnXp, btnYm, btnYp, btnZm, btnZp;
        private Button btnRxm, btnRxp, btnRym, btnRyp, btnRzm, btnRzp;

        private Inventor.UserCoordinateSystem _ucs;
        private Inventor.UnitsOfMeasure _uom;

        public UcsForm(Inventor.Application app)
        {
            _app = app;
            Text = "UCS Inspector – Camarada Escate";
            Width = 700; Height = 440; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false;
            KeyPreview = true; // hotkeys

            // Filtro + combo + pick
            txtFilter = new System.Windows.Forms.TextBox { Left = 12, Top = 12, Width = 220 };
            txtFilter.ForeColor = DrawColor.Gray; txtFilter.Text = "filtrar por nombre…";
            txtFilter.GotFocus += (s, e) => { if (txtFilter.ForeColor == DrawColor.Gray) { txtFilter.Text = ""; txtFilter.ForeColor = DrawColor.Black; } };
            txtFilter.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(txtFilter.Text)) { txtFilter.ForeColor = DrawColor.Gray; txtFilter.Text = "filtrar por nombre…"; } };
            txtFilter.TextChanged += (s, e) => { if (txtFilter.ForeColor != DrawColor.Gray) ApplyFilter(); };

            cboUcs = new ComboBox { Left = 240, Top = 12, Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };
            btnPick = new Button { Left = 530, Top = 12, Width = 140, Text = "Pick en pantalla" };
            btnPick.Click += (s, e) => TryPickUcs();
            Controls.Add(txtFilter); Controls.Add(cboUcs); Controls.Add(btnPick);

            // Filas offsets/ángulos
            MakeRow("X Offset", 55, out lblXo, out txtXo);
            MakeRow("Y Offset", 85, out lblYo, out txtYo);
            MakeRow("Z Offset", 115, out lblZo, out txtZo);
            MakeRow("X Angle", 165, out lblXa, out txtXa);
            MakeRow("Y Angle", 195, out lblYa, out txtYa);
            MakeRow("Z Angle", 225, out lblZa, out txtZa);

            // Nudge buttons a la derecha de cada fila
            AddNudgePair(555, 52, "-X", "+X",
                () => Nudge((Inventor.Parameter)txtXo.Tag, -(double)numStepMm.Value, false),
                () => Nudge((Inventor.Parameter)txtXo.Tag, (double)numStepMm.Value, false), out btnXm, out btnXp);

            AddNudgePair(555, 82, "-Y", "+Y",
                () => Nudge((Inventor.Parameter)txtYo.Tag, -(double)numStepMm.Value, false),
                () => Nudge((Inventor.Parameter)txtYo.Tag, (double)numStepMm.Value, false), out btnYm, out btnYp);

            AddNudgePair(555, 112, "-Z", "+Z",
                () => Nudge((Inventor.Parameter)txtZo.Tag, -(double)numStepMm.Value, false),
                () => Nudge((Inventor.Parameter)txtZo.Tag, (double)numStepMm.Value, false), out btnZm, out btnZp);

            AddNudgePair(555, 162, "-RX", "+RX",
                () => Nudge((Inventor.Parameter)txtXa.Tag, -(double)numStepDeg.Value, true),
                () => Nudge((Inventor.Parameter)txtXa.Tag, (double)numStepDeg.Value, true), out btnRxm, out btnRxp);

            AddNudgePair(555, 192, "-RY", "+RY",
                () => Nudge((Inventor.Parameter)txtYa.Tag, -(double)numStepDeg.Value, true),
                () => Nudge((Inventor.Parameter)txtYa.Tag, (double)numStepDeg.Value, true), out btnRym, out btnRyp);

            AddNudgePair(555, 222, "-RZ", "+RZ",
                () => Nudge((Inventor.Parameter)txtZa.Tag, -(double)numStepDeg.Value, true),
                () => Nudge((Inventor.Parameter)txtZa.Tag, (double)numStepDeg.Value, true), out btnRzm, out btnRzp);

            // Pasos y acciones
            numStepMm = new NumericUpDown { Left = 12, Top = 270, Width = 90, DecimalPlaces = 3, Minimum = 0.001M, Maximum = 10000M, Increment = 0.1M, Value = 1.000M };
            numStepDeg = new NumericUpDown { Left = 110, Top = 270, Width = 90, DecimalPlaces = 2, Minimum = 0.01M, Maximum = 360M, Increment = 0.1M, Value = 1.00M };
            var lblMm = new Label { Left = 12, Top = 250, Width = 90, Text = "Paso mm" };
            var lblDeg = new Label { Left = 110, Top = 250, Width = 90, Text = "Paso °" };
            Controls.Add(lblMm); Controls.Add(lblDeg); Controls.Add(numStepMm); Controls.Add(numStepDeg);

            btnCenter = new Button { Left = 210, Top = 270, Width = 90, Text = "Centrar" };
            btnCopy = new Button { Left = 305, Top = 270, Width = 90, Text = "Copiar" };
            btnPaste = new Button { Left = 400, Top = 270, Width = 90, Text = "Pegar" };
            btnNext = new Button { Left = 495, Top = 270, Width = 90, Text = "Siguiente" };
            btnCenter.Click += (s, e) => CenterOnUcs(_ucs);
            btnCopy.Click += (s, e) => _clipboard = CopyUcs();
            btnPaste.Click += (s, e) => { if (_clipboard.HasValue) PasteUcs(_clipboard.Value); };
            btnNext.Click += (s, e) => NextUcs();
            Controls.AddRange(new Control[] { btnCenter, btnCopy, btnPaste, btnNext });

            btnRefresh = new Button { Left = 12, Top = 315, Width = 140, Text = "Actualizar" };
            btnApply = new Button { Left = 530, Top = 315, Width = 140, Text = "Aplicar" };
            btnRefresh.Click += (s, e) => LoadValues();
            btnApply.Click += (s, e) => ApplyChanges();
            Controls.Add(btnRefresh); Controls.Add(btnApply);

            LoadUcsList();
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

        private void LoadUcsList()
        {
            cboUcs.Items.Clear();
            _allUcs.Clear();
            _uom = ActiveDoc != null ? ActiveDoc.UnitsOfMeasure : null;

            Inventor.UserCoordinateSystems col = null;
            if (ActiveDoc is Inventor.PartDocument)
                col = ((Inventor.PartDocument)ActiveDoc).ComponentDefinition.UserCoordinateSystems;
            else if (ActiveDoc is Inventor.AssemblyDocument)
                col = ((Inventor.AssemblyDocument)ActiveDoc).ComponentDefinition.UserCoordinateSystems;

            if (col == null || col.Count == 0)
            {
                MessageBox.Show("No hay UCS en el documento actual.");
                return;
            }

            for (int i = 1; i <= col.Count; i++) _allUcs.Add(col[i]);

            cboUcs.DisplayMember = "Name";
            cboUcs.SelectedIndexChanged -= CboUcs_SelectedIndexChanged;
            cboUcs.SelectedIndexChanged += CboUcs_SelectedIndexChanged;

            ApplyFilter();
        }

        private void CboUcs_SelectedIndexChanged(object sender, EventArgs e)
        {
            _ucs = (Inventor.UserCoordinateSystem)cboUcs.SelectedItem;
            LoadValues();
        }

        private void ApplyFilter()
        {
            var f = (txtFilter.Text ?? "").Trim();
            var filtered = (txtFilter.ForeColor == DrawColor.Gray || string.IsNullOrEmpty(f))
                ? _allUcs
                : _allUcs.Where(u => u.Name.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

            cboUcs.Items.Clear();
            foreach (var u in filtered) cboUcs.Items.Add(u);
            if (cboUcs.Items.Count > 0) cboUcs.SelectedIndex = 0;
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

            lblXo.Text = ToStr(px);
            lblYo.Text = ToStr(py);
            lblZo.Text = ToStr(pz);
            lblXa.Text = ToStr(ax);
            lblYa.Text = ToStr(ay);
            lblZa.Text = ToStr(az);

            txtXo.Tag = px; txtYo.Tag = py; txtZo.Tag = pz;
            txtXa.Tag = ax; txtYa.Tag = ay; txtZa.Tag = az;
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
            try
            {
                ApplyOne((Inventor.Parameter)txtXo.Tag, txtXo.Text, false);
                ApplyOne((Inventor.Parameter)txtYo.Tag, txtYo.Text, false);
                ApplyOne((Inventor.Parameter)txtZo.Tag, txtZo.Text, false);
                ApplyOne((Inventor.Parameter)txtXa.Tag, txtXa.Text, true);
                ApplyOne((Inventor.Parameter)txtYa.Tag, txtYa.Text, true);
                ApplyOne((Inventor.Parameter)txtZa.Tag, txtZa.Text, true);

                ActiveDoc.Update();
                LoadValues();
                MessageBox.Show("Cambios aplicados.");
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
        }

        private void ApplyOne(Inventor.Parameter p, string expr, bool isAngle)
        {
            if (p == null) return;
            expr = (expr ?? "").Trim();
            if (string.IsNullOrWhiteSpace(expr)) return;

            string unitsStr = GetUnits(p);
            var unitsEnum = UnitEnum(unitsStr, isAngle);

            char op = expr[0];
            double v = GetVal(p);

            if (op == '+' || op == '-')
            {
                double delta = Convert.ToDouble(_uom.GetValueFromExpression(expr.Substring(1).Trim(), unitsEnum));
                SetVal(p, (op == '+') ? v + delta : v - delta);
            }
            else if (op == '*' || op == '/')
            {
                var num = double.Parse(expr.Substring(1).Trim(), CultureInfo.InvariantCulture);
                SetVal(p, (op == '*') ? v * num : v / num);
            }
            else
            {
                double abs = Convert.ToDouble(_uom.GetValueFromExpression(expr, unitsEnum));
                SetVal(p, abs);
            }
        }

        // --- NUDGES (pasos) + HOTKEYS + UNDO ---

        private void Nudge(Inventor.Parameter p, double delta, bool isAngle)
        {
            if (p == null) return;

            string unitsStr = GetUnits(p);
            var unitsEnum = UnitEnum(unitsStr, isAngle);

            var tm = _app.TransactionManager;
            Inventor.Transaction tx = null;
            try
            {
                tx = tm.StartTransaction((Inventor._Document)ActiveDoc, "UCS nudge");
                double v = GetVal(p);
                double inc = Convert.ToDouble(_uom.GetValueFromExpression(
                    delta.ToString(CultureInfo.InvariantCulture), unitsEnum));
                SetVal(p, v + inc);
                ActiveDoc.Update();
                tx.End(); // un solo paso en Undo
                LoadValues();
            }
            catch
            {
                if (tx != null) tx.Abort();
                throw;
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            double stepMM = numStepMm != null ? (double)numStepMm.Value : 1.0;
            double stepDeg = numStepDeg != null ? (double)numStepDeg.Value : 1.0;
            if (_ucs == null) return base.ProcessCmdKey(ref msg, keyData);

            var px = (Inventor.Parameter)txtXo.Tag; var py = (Inventor.Parameter)txtYo.Tag; var pz = (Inventor.Parameter)txtZo.Tag;
            var ax = (Inventor.Parameter)txtXa.Tag; var ay = (Inventor.Parameter)txtYa.Tag; var az = (Inventor.Parameter)txtZa.Tag;

            switch (keyData)
            {
                case (Keys.Alt | Keys.Q): Nudge(px, -stepMM, false); return true; // X-
                case (Keys.Alt | Keys.W): Nudge(px, stepMM, false); return true;  // X+
                case (Keys.Alt | Keys.A): Nudge(py, -stepMM, false); return true; // Y-
                case (Keys.Alt | Keys.S): Nudge(py, stepMM, false); return true;  // Y+
                case (Keys.Alt | Keys.Z): Nudge(pz, -stepMM, false); return true; // Z-
                case (Keys.Alt | Keys.X): Nudge(pz, stepMM, false); return true;  // Z+

                case (Keys.Alt | Keys.D1): Nudge(ax, stepDeg, true); return true; // RX+
                case (Keys.Alt | Keys.D2): Nudge(ax, -stepDeg, true); return true; // RX-
                case (Keys.Alt | Keys.D3): Nudge(ay, stepDeg, true); return true; // RY+
                case (Keys.Alt | Keys.D4): Nudge(ay, -stepDeg, true); return true; // RY-
                case (Keys.Alt | Keys.D5): Nudge(az, stepDeg, true); return true; // RZ+
                case (Keys.Alt | Keys.D6): Nudge(az, -stepDeg, true); return true; // RZ-

                case Keys.Enter: NextUcs(); return true; // recorrer
            }
            return base.ProcessCmdKey(ref msg, keyData);
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
            return (V((Inventor.Parameter)txtXo.Tag), V((Inventor.Parameter)txtYo.Tag), V((Inventor.Parameter)txtZo.Tag),
                    V((Inventor.Parameter)txtXa.Tag), V((Inventor.Parameter)txtYa.Tag), V((Inventor.Parameter)txtZa.Tag));
        }

        private void PasteUcs((double xo, double yo, double zo, double xa, double ya, double za) v)
        {
            void Set(Inventor.Parameter p, double val, bool ang)
            {
                if (p == null) return;
                string u = GetUnits(p);
                var ut = UnitEnum(u, ang);
                p.Value = _uom.GetValueFromExpression(val.ToString(CultureInfo.InvariantCulture), ut);
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
                object u = p.get_Units(); // descriptor de acceso COM
                if (u is string) return (string)u;
                return u != null ? u.ToString() : "";
            }
            catch { return ""; }
        }

        private double GetVal(Inventor.Parameter p)
        {
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
    }
}
