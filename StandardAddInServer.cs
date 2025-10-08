using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Inventor;
using System.Web.Script.Serialization; // referencia en el proyecto: System.Web.Extensions
using System.Windows.Forms;            // Clipboard


// ALIAS seguros con Inventor
using SysEnv = System.Environment;
using SysPath = System.IO.Path;
using SysFile = System.IO.File;
using SysDir = System.IO.Directory;

namespace UcsInspectorperu
{
    [ComVisible(true)]
    [Guid("a8c2eab2-d332-4188-bf7d-b9a07768fe66")]
    public class StandardAddInServer : ApplicationAddInServer
    {
        // ---------- LOG: %LOCALAPPDATA%\UcsInspectorperu\addin.log (fallback Roaming) ----------
        private static void Log(string msg)
        {
            var local = SysPath.Combine(SysEnv.GetFolderPath(SysEnv.SpecialFolder.LocalApplicationData), "UcsInspectorperu", "addin.log");
            var roam = SysPath.Combine(SysEnv.GetFolderPath(SysEnv.SpecialFolder.ApplicationData), "UcsInspectorperu", "addin.log");
            string[] paths = new[] { local, roam };

            foreach (var p in paths)
            {
                try
                {
                    var dir = SysPath.GetDirectoryName(p);
                    if (!string.IsNullOrEmpty(dir)) SysDir.CreateDirectory(dir);
                    SysFile.AppendAllText(p, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + msg + SysEnv.NewLine);
                    return;
                }
                catch { /* prueba la siguiente */ }
            }
        }

        // ---------- RESOLVE del Interop + FirstChance (sin pop-ups) ----------
        static StandardAddInServer()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                try
                {
                    if (args.Name.StartsWith("Autodesk.Inventor.Interop", StringComparison.OrdinalIgnoreCase))
                    {
                        string pf = SysEnv.GetFolderPath(SysEnv.SpecialFolder.ProgramFiles);
                        string[] guesses =
                        {
                            SysPath.Combine(pf, @"Autodesk\Inventor 2026\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll"),
                            SysPath.Combine(pf, @"Autodesk\Inventor 2026\Bin\Autodesk.Inventor.Interop.dll")
                        };
                        foreach (var g in guesses)
                            if (SysFile.Exists(g)) return Assembly.LoadFrom(g);
                    }
                }
                catch { }
                return null;
            };

            AppDomain.CurrentDomain.FirstChanceException += (s, e) =>
            {
                Log("FirstChance: " + e.Exception.GetType().FullName + " - " + e.Exception.Message);
            };

            Log("static ctor reached");
        }

        private Inventor.Application _app;
        private ButtonDefinition _btn;
        private readonly string _clientId = typeof(StandardAddInServer).GUID.ToString("B");

        public StandardAddInServer()
        {
            Log("ctor StandardAddInServer()");
        }

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                Log("Activate() start. firstTime=" + firstTime);
                _app = addInSiteObject.Application;
                Log("Assembly.Location = " + typeof(StandardAddInServer).Assembly.Location);

                // --- Iconos embebidos (opcionales). Si no se encuentran, sigue sin iconos.
                var smallIco = LoadIcon(".res.ucs16.png");
                var largeIco = LoadIcon(".res.ucs32.png");

                var defs = _app.CommandManager.ControlDefinitions;

                if (smallIco != null && largeIco != null)
                {
                    _btn = defs.AddButtonDefinition(
                        "UCS Inspector",
                        "UcsInspectorperu:UcsInspectorBtn",
                        CommandTypesEnum.kNonShapeEditCmdType,
                        _clientId,
                        "Ver y editar offsets/ángulos de un UCS",
                        "Selecciona un UCS y aplica +, -, * o / a sus parámetros",
                        smallIco, largeIco, ButtonDisplayEnum.kDisplayTextInLearningMode);
                }
                else
                {
                    _btn = defs.AddButtonDefinition(
                        "UCS Inspector",
                        "UcsInspectorperu:UcsInspectorBtn",
                        CommandTypesEnum.kNonShapeEditCmdType,
                        _clientId,
                        "Ver y editar offsets/ángulos de un UCS",
                        "Selecciona un UCS y aplica +, -, * o / a sus parámetros");
                }

                _btn.OnExecute += (Context) =>
                {
                    try
                    {
                        var owner = new WindowWrapper(new IntPtr(_app.MainFrameHWND));
                        var frm = new Ui.UcsForm(_app)
                        {
                            ShowInTaskbar = false,
                            StartPosition = FormStartPosition.CenterParent
                        };
                        frm.Show(owner);
                    }
                    catch (Exception ex) { Log("OnExecute error: " + ex); MessageBox.Show(ex.ToString(), "UcsInspectorperu"); }
                };

                TryAddToRibbon("Part", "id_TabTools", "UcsInspectorperu:PanelPart");
                TryAddToRibbon("Assembly", "id_TabTools", "UcsInspectorperu:PanelAsm");
                TryAddToRibbon("ZeroDoc", "id_TabToolsZeroDoc", "UcsInspectorperu:PanelZero");
                TryAddToRibbon("Drawing", "id_TabTools", "UcsInspectorperu:PanelDrw");

                // Hotkey global (Ctrl+Shift+U por defecto; configurable en settings.json)
                SetupGlobalHotkey();

                Log("Activate() OK");
            }
            catch (Exception ex)
            {
                Log("Activate() FAILED: " + ex);
                MessageBox.Show("UCS Inspector no pudo cargarse:\n\n" + ex, "UcsInspectorperu",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        public void Deactivate()
        {
            Log("Deactivate()");
            try { HotkeyService.Unregister(); } catch { }
            try { if (_btn != null) Marshal.ReleaseComObject(_btn); } catch { }
            try { if (_app != null) Marshal.ReleaseComObject(_app); } catch { }
            _btn = null; _app = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID) { }
        public object Automation { get { return null; } }

        // Wrapper simple para pasar HWND como owner de WinForms
        private sealed class WindowWrapper : IWin32Window
        {
            private readonly IntPtr _hwnd;
            public WindowWrapper(IntPtr handle) { _hwnd = handle; }
            public IntPtr Handle { get { return _hwnd; } }
        }

        // ---------- Carga de iconos embebidos como stdole.IPictureDisp ----------
        private static stdole.IPictureDisp LoadIcon(string resourceNameEndsWith)
        {
            try
            {
                var asm = typeof(StandardAddInServer).Assembly;
                string name = asm.GetManifestResourceNames()
                                 .FirstOrDefault(n => n.EndsWith(resourceNameEndsWith, StringComparison.OrdinalIgnoreCase));
                if (name == null) { Log("Icono no encontrado (sufijo): " + resourceNameEndsWith); return null; }

                using (var s = asm.GetManifestResourceStream(name))
                using (var img = System.Drawing.Image.FromStream(s))
                using (var bmp24 = new System.Drawing.Bitmap(img.Width, img.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
                using (var g = System.Drawing.Graphics.FromImage(bmp24))
                {
                    g.Clear(System.Drawing.Color.Transparent);
                    g.DrawImage(img, 0, 0, img.Width, img.Height); // sin alpha al IPictureDisp
                    return (stdole.IPictureDisp)AxHostWrapper.GetIPictureDispFromPicture(bmp24);
                }
            }
            catch (Exception ex)
            {
                Log("LoadIcon fallo: " + ex.Message);
                return null;
            }
        }


        private class AxHostWrapper : AxHost
        {
            private AxHostWrapper() : base("") { }
            public static object GetIPictureDispFromPicture(System.Drawing.Image img)
            {
                return AxHost.GetIPictureDispFromPicture(img);
            }
        }

        // al final del archivo, dentro del namespace UcsInspectorperu (puede ser clase interna)
        // Clase auxiliar para registrar un atajo global del sistema (Ctrl+Shift+U por defecto).
        internal static class HotkeyService
        {
            // Win32
            [System.Runtime.InteropServices.DllImport("user32.dll")]
            private static extern bool RegisterHotKey(System.IntPtr hWnd, int id, uint fsModifiers, uint vk);

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            private static extern bool UnregisterHotKey(System.IntPtr hWnd, int id);

            private const int WM_HOTKEY = 0x0312;
            private const uint MOD_ALT = 0x0001;
            private const uint MOD_CONTROL = 0x0002;
            private const uint MOD_SHIFT = 0x0004;

            // Ventana oculta para recibir WM_HOTKEY
            private sealed class MsgWnd : NativeWindow
            {
                private readonly ButtonDefinition _btn;

                public MsgWnd(System.IntPtr parent, ButtonDefinition btn)
                {
                    _btn = btn;
                    var cp = new CreateParams();
                    cp.Parent = parent;             // colgarla del frame principal de Inventor
                    CreateHandle(cp);
                }

                protected override void WndProc(ref Message m)
                {
                    if (m.Msg == WM_HOTKEY && _btn != null)
                    {
                        try { _btn.Execute(); } catch { /* no romper Inventor */ }
                    }
                    base.WndProc(ref m);
                }
            }

            private static MsgWnd _wnd;
            private static int _regId;

            /// <summary>
            /// Registra el atajo global. Ejemplos válidos: "Ctrl+Shift+U", "Alt+U", "Ctrl+I".
            /// </summary>
            public static void Register(Inventor.Application app, ButtonDefinition btn, string hotkeyText)
            {
                try
                {
                    Unregister(); // por si estaba registrado antes

                    uint mods, vk;
                    ParseHotkey(hotkeyText, out mods, out vk);

                    // Crea la ventana receptora y registra el hotkey
                    _wnd = new MsgWnd(new System.IntPtr(app.MainFrameHWND), btn);
                    _regId = 1; // único id
                    RegisterHotKey(_wnd.Handle, _regId, mods, vk);
                }
                catch
                {
                    // Silencioso: si falla, el add-in sigue funcionando sin atajo global
                }
            }

            /// <summary>
            /// Libera el atajo global.
            /// </summary>
            public static void Unregister()
            {
                try
                {
                    if (_wnd != null && _regId != 0)
                        UnregisterHotKey(_wnd.Handle, _regId);
                }
                catch { }
                finally
                {
                    try { if (_wnd != null) _wnd.DestroyHandle(); } catch { }
                    _wnd = null;
                    _regId = 0;
                }
            }

            /// <summary>
            /// Convierte una cadena tipo "Ctrl+Shift+U" en modificadores y tecla virtual.
            /// </summary>
            private static void ParseHotkey(string s, out uint mods, out uint vk)
            {
                mods = 0;
                vk = (uint)Keys.U; // por defecto

                try
                {
                    string txt = (s ?? "").Trim().ToUpperInvariant();
                    if (txt.Length == 0) { mods = MOD_CONTROL | MOD_SHIFT; return; }

                    string[] parts = txt.Split(new char[] { '+', '-', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var p0 in parts)
                    {
                        string p = p0.Trim();
                        if (p == "CTRL" || p == "CONTROL") { mods |= MOD_CONTROL; continue; }
                        if (p == "SHIFT") { mods |= MOD_SHIFT; continue; }
                        if (p == "ALT") { mods |= MOD_ALT; continue; }

                        // Tecla principal
                        Keys k;
                        if (Enum.TryParse<Keys>(p, true, out k))
                        {
                            vk = (uint)k;
                        }
                        else if (p.Length == 1)
                        {
                            vk = (uint)char.ToUpperInvariant(p[0]);
                        }
                    }

                    if (mods == 0) mods = MOD_CONTROL | MOD_SHIFT; // seguridad
                }
                catch
                {
                    mods = MOD_CONTROL | MOD_SHIFT; vk = (uint)Keys.U;
                }
            }
        }


      
        // Añade esto dentro de StandardAddInServer (misma clase)

        private void TryAddToRibbon(string ribbonName, string tabId, string panelInternalName)
        {
            try
            {
                var rb = _app.UserInterfaceManager.Ribbons[ribbonName];
                var tab = rb.RibbonTabs[tabId];

                RibbonPanel panel;
                try { panel = tab.RibbonPanels["id_PanelPgmAddIns"]; } // panel integrado de Add-Ins
                catch { panel = tab.RibbonPanels.Add("Add-Ins", panelInternalName, _clientId); }

                bool already = false;
                foreach (CommandControl c in panel.CommandControls)
                    if (c.InternalName == _btn.InternalName) { already = true; break; }

                if (!already) panel.CommandControls.AddButton(_btn, true);

                Log("Ribbon OK: " + ribbonName + "/" + tabId);
            }
            catch (Exception ex)
            {
                Log("Ribbon skip " + ribbonName + "/" + tabId + ": " + ex.Message);
            }
        }

        private void SetupGlobalHotkey()
        {
            try
            {
                bool enable = true;
                string hotkey = "Ctrl+Shift+U";

                // %LOCALAPPDATA%\UcsInspectorperu\settings.json
                string cfgPath = SysPath.Combine(
                    SysEnv.GetFolderPath(SysEnv.SpecialFolder.LocalApplicationData),
                    "UcsInspectorperu", "settings.json");

                try
                {
                    if (SysFile.Exists(cfgPath))
                    {
                        var js = SysFile.ReadAllText(cfgPath);

                        if (js.IndexOf("\"UseGlobalHotkey\": false", StringComparison.OrdinalIgnoreCase) >= 0)
                            enable = false;

                        int k = js.IndexOf("\"GlobalHotkey\"", StringComparison.OrdinalIgnoreCase);
                        if (k >= 0)
                        {
                            int q1 = js.IndexOf('"', k + 14);
                            q1 = js.IndexOf('"', q1 + 1);
                            int q2 = js.IndexOf('"', q1 + 1);
                            if (q1 > 0 && q2 > q1) hotkey = js.Substring(q1 + 1, q2 - q1 - 1);
                        }
                    }
                }
                catch { /* settings corrupto → usa defaults */ }

                if (enable)
                {
                    HotkeyService.Register(_app, _btn, hotkey);
                    Log("Global hotkey activo: " + hotkey);
                }
                else
                {
                    Log("Global hotkey desactivado por settings.");
                }
            }
            catch (Exception ex)
            {
                Log("SetupGlobalHotkey() fallo: " + ex.Message);
            }
        }



    }
}
