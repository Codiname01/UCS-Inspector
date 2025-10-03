using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;    // <-- necesario para el Clipboard JSON
using System.Windows.Forms;
using Inventor;

using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.IO;


// Alias seguros
using DrawColor = System.Drawing.Color;
using WinTextBox = System.Windows.Forms.TextBox;
using SysEnv = System.Environment;
using SysPath = System.IO.Path;


namespace UcsInspectorperu.Ui
{
    public class UcsForm : Form
    {
        private readonly Inventor.Application _app;
        private ComboBox cboUcs;
        private Label lblXo, lblYo, lblZo, lblXa, lblYa, lblZa;
        private WinTextBox txtXo, txtYo, txtZo, txtXa, txtYa, txtZa;
        private Button btnApply, btnRefresh, btnPick;
        private Button btnHelp;
        private (double xo, double yo, double zo, double xa, double ya, double za)? _clipboard;
        private CheckBox chkDelta;
        private Button btnViewUcs;





        // NUEVO: UI y estado
        private NumericUpDown numStepMm, numStepDeg;
        private Button btnCenter, btnCopy, btnPaste, btnNext;
        private System.Windows.Forms.TextBox txtFilter;
        private readonly List<Inventor.UserCoordinateSystem> _allUcs = new List<Inventor.UserCoordinateSystem>();
        


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
            // en btnPaste.Click:
            (double xo, double yo, double zo, double xa, double ya, double za) data;
            ClipPacket raw;
            if (TryGetFromClipboard(out data, out raw))
            {
                // Si hay expresiones, escríbelas literal; si no, pega valores
                Action<WinTextBox, string, bool, double> apply = (tb, exp, ang, val) =>
                {
                    var p = (Inventor.Parameter)tb.Tag;
                    if (p == null) return;
                    if (!string.IsNullOrWhiteSpace(exp)) { p.Expression = exp.StartsWith("=") ? exp.Substring(1).Trim() : exp; }
                    else { SetVal(p, val, ang); }
                };
                apply(txtXo, raw.xoExp, false, data.xo);
                apply(txtYo, raw.yoExp, false, data.yo);
                apply(txtZo, raw.zoExp, false, data.zo);
                apply(txtXa, raw.xaExp, true, data.xa);
                apply(txtYa, raw.yaExp, true, data.ya);
                apply(txtZa, raw.zaExp, true, data.za);

                ActiveDoc.Update(); LoadValues();
            }
            else if (_clipboard.HasValue) { PasteUcs(_clipboard.Value); }
            else MessageBox.Show("No hay datos de UCS en el portapapeles.");

            btnNext.Click += (s, e) => NextUcs();
            Controls.AddRange(new Control[] { btnCenter, btnCopy, btnPaste, btnNext });

            btnRefresh = new Button { Left = 12, Top = 315, Width = 140, Text = "Actualizar" };
            btnApply = new Button { Left = 530, Top = 315, Width = 140, Text = "Aplicar" };
            btnHelp = new Button { Left = 405, Top = 315, Width = 115, Text = "Ayuda (F1)" };
            btnHelp.Click += (s, e) => ShowHelp();
            Controls.Add(btnHelp);

            btnRefresh.Click += (s, e) => LoadValues();
            btnApply.Click += (s, e) => ApplyChanges();
            Controls.Add(btnRefresh); Controls.Add(btnApply);

            LoadUcsList();

            var btnCopyExpr = new Button { Left = 305, Top = 315, Width = 90, Text = "Copiar expr" };
            btnCopyExpr.Click += (s, e) => CopyUcsExpressionsToClipboard();
            Controls.Add(btnCopyExpr);

            // Dentro del constructor, una sola vez:
            btnPaste.Click += (s, e) =>
            {
                // NOMBRES DIFERENTES para evitar CS0136
                (double xo, double yo, double zo, double xa, double ya, double za) clipVals;
                ClipPacket clipPkt;

                if (TryGetFromClipboard(out clipVals, out clipPkt))
                {
                    if (chkDelta.Checked)
                    {
                        PasteAsDelta(clipVals);
                    }
                    else
                    {
                        // Si hay expresiones en el portapapeles, respétalas; si no, pega valores
                        Action<WinTextBox, string, bool, double> apply = (tb, expr, isAng, val) =>
                        {
                            var p = (Inventor.Parameter)tb.Tag;
                            if (p == null) return;
                            if (clipPkt != null && !string.IsNullOrWhiteSpace(expr))
                                p.Expression = expr.StartsWith("=") ? expr.Substring(1).Trim() : expr;
                            else
                                SetVal(p, val, isAng);
                        };

                        apply(txtXo, clipPkt != null ? clipPkt.xoExp : null, false, clipVals.xo);
                        apply(txtYo, clipPkt != null ? clipPkt.yoExp : null, false, clipVals.yo);
                        apply(txtZo, clipPkt != null ? clipPkt.zoExp : null, false, clipVals.zo);
                        apply(txtXa, clipPkt != null ? clipPkt.xaExp : null, true, clipVals.xa);
                        apply(txtYa, clipPkt != null ? clipPkt.yaExp : null, true, clipVals.ya);
                        apply(txtZa, clipPkt != null ? clipPkt.zaExp : null, true, clipVals.za);

                        ActiveDoc.Update();
                        LoadValues();
                    }
                }
                else if (_clipboard.HasValue)
                {
                    PasteUcs(_clipboard.Value);
                }
                else
                {
                    MessageBox.Show("No hay datos de UCS en el portapapeles.");
                }
            };

            // Botón: Vista = UCS (Z arriba, X a la derecha)
            btnViewUcs = new Button { Left = 210, Top = 345, Width = 120, Text = "Vista = UCS" };
            btnViewUcs.Click += (s, e) => AlignViewToUcs(_ucs);
            Controls.Add(btnViewUcs);




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




        [DataContract]
        public class ClipPacket
        {
            [DataMember] public string kind;   // "UCS-vals" o "UCS-expr"
            [DataMember] public int ver;    // 2

            // valores con unidades tipo "92 mm" / "10 deg"
            [DataMember] public string xo;
            [DataMember] public string yo;
            [DataMember] public string zo;
            [DataMember] public string xa;
            [DataMember] public string ya;
            [DataMember] public string za;


            // si copiaste expresiones, van aquí (opcional)
            [DataMember] public string xoExp;
            [DataMember] public string yoExp;
            [DataMember] public string zoExp;
            [DataMember] public string xaExp;
            [DataMember] public string yaExp;
            [DataMember] public string zaExp;
        }

        private static string ToJson<T>(T obj)
        {
            try
            {
                var ser = new DataContractJsonSerializer(typeof(T));
                using (var ms = new MemoryStream())
                {
                    ser.WriteObject(ms, obj);
                    return Encoding.UTF8.GetString(ms.ToArray());
                }
            }
            catch { return null; }
        }

        private static bool TryFromJson<T>(string s, out T obj)
        {
            obj = default(T);
            try
            {
                var ser = new DataContractJsonSerializer(typeof(T));
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(s ?? "")))
                {
                    obj = (T)ser.ReadObject(ms);
                    return true;
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

            // 1) Expresión cruda (como en Inventor): pásala directo a Expression
            if (LooksLikeRawExpression(expr))
            {
                if (expr.StartsWith("=")) expr = expr.Substring(1).Trim();
                p.Expression = expr;            // deja que Inventor la evalúe
                return;
            }

            // 2) Operaciones relativas
            string unitsStr = GetUnits(p);
            var unitsEnum = UnitEnum(unitsStr, isAngle);

            char op = expr[0];
            double v = GetVal(p);

            if (op == '+' || op == '-')
            {
                double delta = Convert.ToDouble(_uom.GetValueFromExpression(expr.Substring(1).Trim(), unitsEnum));
                SetVal(p, (op == '+') ? v + delta : v - delta, isAngle);
            }
            else if (op == '*' || op == '/')
            {
                double factor;
                if (!double.TryParse(expr.Substring(1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out factor))
                    throw new FormatException("Factor inválido: " + expr);

                SetVal(p, (op == '*') ? v * factor : v / factor, isAngle);
            }
            else
            {
                // 3) Absoluto numérico con unidades
                double abs = Convert.ToDouble(_uom.GetValueFromExpression(expr, unitsEnum));
                SetVal(p, abs, isAngle);
            }
        }

        // --- NUDGES (pasos) + HOTKEYS + UNDO ---

        private void Nudge(Inventor.Parameter p, double delta, bool isAngle)
        {
            if (p == null) return;
            WithFast(() =>
            {
                string unitsStr = GetUnits(p);
                var unitsEnum = UnitEnum(unitsStr, isAngle);

                var tm = _app.TransactionManager; Inventor.Transaction tx = null;
                try
                {
                    tx = tm.StartTransaction((Inventor._Document)ActiveDoc, "UCS nudge");
                    double v = GetVal(p);
                    double inc = Convert.ToDouble(_uom.GetValueFromExpression(
                        delta.ToString(CultureInfo.InvariantCulture), unitsEnum));
                    SetVal(p, v + inc, isAngle);
                    ActiveDoc.Update();
                    tx.End();
                    LoadValues();
                }
                catch { if (tx != null) tx.Abort(); throw; }
            });
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
                var s = (Clipboard.GetText() ?? "").Trim();
                if (s.Length == 0) return false;

                // ¿Es JSON nuestro?
                if (s.StartsWith("{") && s.IndexOf("\"kind\"", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ClipPacket pkt;
                    if (TryFromJson<ClipPacket>(s, out pkt) && pkt != null && !string.IsNullOrEmpty(pkt.kind))
                    {
                        raw = pkt;

                        // Convierte los strings con unidades a valores numéricos (mm/deg)
                        double Val(string t, bool ang)
                        {
                            if (string.IsNullOrWhiteSpace(t)) return 0.0;
                            var ut = ang ? Inventor.UnitsTypeEnum.kDegreeAngleUnits
                                         : Inventor.UnitsTypeEnum.kMillimeterLengthUnits;
                            return Convert.ToDouble(_uom.GetValueFromExpression(t, ut));
                        }

                        v = (Val(pkt.xo, false), Val(pkt.yo, false), Val(pkt.zo, false),
                             Val(pkt.xa, true), Val(pkt.ya, true), Val(pkt.za, true));
                        return true;
                    }
                }

                // No es nuestro JSON → no tocar
                return false;
            }
            catch
            {
                // nunca interrumpas la carga del add-in por esto
                return false;
            }
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
        private void SetVal(Inventor.Parameter p, double v, bool isAngle)
        {
            if (p == null) return;
            try { p.Value = v; }
            catch (System.Runtime.InteropServices.COMException)
            {
                var ut = UnitEnum(GetUnits(p), isAngle);
                string exprText = _uom.GetStringFromValue(v, ut); // ej. "50 mm" / "10 deg"
                p.Expression = exprText;
            }
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

        private void CopyUcsExpressionsToClipboard()
        {
            var pack = new ClipPacket { kind = "UCS-expr", ver = 2 };

            string Val(Inventor.Parameter p, bool ang)
            {
                if (p == null) return null;
                var ut = UnitEnum(GetUnits(p), ang);
                return _uom.GetStringFromValue(GetVal(p), ut); // "92 mm", "10 deg", etc.
            }

            string Exp(Inventor.Parameter p)
            {
                if (p == null) return null;
                try { return p.Expression; } catch { return null; }
            }

            // valores con unidades (compat)
            pack.xo = Val((Inventor.Parameter)txtXo.Tag, false);
            pack.yo = Val((Inventor.Parameter)txtYo.Tag, false);
            pack.zo = Val((Inventor.Parameter)txtZo.Tag, false);
            pack.xa = Val((Inventor.Parameter)txtXa.Tag, true);
            pack.ya = Val((Inventor.Parameter)txtYa.Tag, true);
            pack.za = Val((Inventor.Parameter)txtZa.Tag, true);

            // expresiones literales
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










        private void PasteAsDelta((double xo, double yo, double zo, double xa, double ya, double za) d)
        {
            WithFast(() =>
            {
                var tm = _app.TransactionManager; Inventor.Transaction tx = null;
                try
                {
                    tx = tm.StartTransaction((Inventor._Document)ActiveDoc, "UCS paste delta");

                    void Add(WinTextBox tb, double dv, bool ang)
                    {
                        var p = (Inventor.Parameter)tb.Tag; if (p == null || IsReadOnly(p)) return;
                        SetVal(p, GetVal(p) + dv, ang);
                    }
                    Add(txtXo, d.xo, false); Add(txtYo, d.yo, false); Add(txtZo, d.zo, false);
                    Add(txtXa, d.xa, true); Add(txtYa, d.ya, true); Add(txtZa, d.za, true);

                    ActiveDoc.Update(); tx.End(); LoadValues();
                }
                catch { if (tx != null) tx.Abort(); throw; }
            });
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





    }
}
