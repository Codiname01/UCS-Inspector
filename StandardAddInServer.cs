using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;
using Inventor;

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

                // Intentar crear atajo Ctrl+Shift+U (si la API está disponible)
                TryEnsureShortcut();

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

        // --- Crear atajo por reflexión si Inventor expone KeyboardShortcuts
        private void TryEnsureShortcut()
        {
            try
            {
                var uim = _app.UserInterfaceManager;
                var ksObj = uim.GetType().InvokeMember("KeyboardShortcuts",
                    BindingFlags.GetProperty, null, uim, new object[0]);

                if (ksObj == null) { Log("KeyboardShortcuts no disponible."); return; }

                // ¿ya existe un atajo para nuestro botón?
                foreach (object item in (System.Collections.IEnumerable)ksObj)
                {
                    var def = item.GetType().GetProperty("CommandDefinition")?.GetValue(item, null);
                    var inName = def?.GetType().GetProperty("InternalName")?.GetValue(def, null) as string;
                    if (string.Equals(inName, _btn.InternalName, StringComparison.OrdinalIgnoreCase))
                        return; // ya está
                }

                // Intentar agregar "Ctrl+Shift+U"
                try
                {
                    ksObj.GetType().InvokeMember("Add",
                        BindingFlags.InvokeMethod, null, ksObj,
                        new object[] { _btn, "Ctrl+Shift+U" });
                    Log("Shortcut creado: Ctrl+Shift+U");
                }
                catch (Exception ex2)
                {
                    Log("No se pudo crear shortcut por API: " + ex2.Message);
                }
            }
            catch (Exception ex) { Log("TryEnsureShortcut() fallo: " + ex.Message); }
        }

        public void Deactivate()
        {
            Log("Deactivate()");
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
                // Buscar por sufijo para no depender del namespace exacto
                string name = asm.GetManifestResourceNames()
                                 .FirstOrDefault(n => n.EndsWith(resourceNameEndsWith, StringComparison.OrdinalIgnoreCase));
                if (name == null) { Log("Icono no encontrado (sufijo): " + resourceNameEndsWith); return null; }

                using (var s = asm.GetManifestResourceStream(name))
                using (var img = System.Drawing.Image.FromStream(s))
                {
                    return (stdole.IPictureDisp)AxHostWrapper.GetIPictureDispFromPicture(img);
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
    }
}
