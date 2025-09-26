using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;

// ALIAS para evitar choques con Inventor.Environment/File/Path
using Env = System.Environment;
using IOPath = System.IO.Path;
using IOFile = System.IO.File;
using IODir = System.IO.Directory;

namespace UcsInspectorperu
{
    [ComVisible(true)]
    [Guid("a8c2eab2-d332-4188-bf7d-b9a07768fe66")]
    public class StandardAddInServer : ApplicationAddInServer
    {
        // ---------- LOG (con varios fallbacks) ----------
        private static void Log(string msg)
        {
            string[] paths = new[]
            {
                IOPath.Combine(Env.GetFolderPath(Env.SpecialFolder.ApplicationData),      "UcsInspectorperu", "addin.log"),
                IOPath.Combine(Env.GetFolderPath(Env.SpecialFolder.LocalApplicationData), "Temp",             "UcsInspectorperu.addin.log"),
                IOPath.Combine(Env.GetEnvironmentVariable("TEMP") ?? IOPath.GetTempPath(),"UcsInspectorperu.addin.log"),
                IOPath.Combine(Env.GetFolderPath(Env.SpecialFolder.MyDocuments),          "UcsInspectorperu", "addin.log"),
            };

            foreach (var p in paths)
            {
                try
                {
                    var dir = IOPath.GetDirectoryName(p);
                    if (!string.IsNullOrEmpty(dir)) IODir.CreateDirectory(dir);
                    IOFile.AppendAllText(p, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + msg + Env.NewLine);
                    return;
                }
                catch { /* prueba la siguiente ruta */ }
            }
        }

        // ---------- RESOLVE del Interop + pop-up temprano ----------
        static StandardAddInServer()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                try
                {
                    if (args.Name.StartsWith("Autodesk.Inventor.Interop", StringComparison.OrdinalIgnoreCase))
                    {
                        string pf = Env.GetFolderPath(Env.SpecialFolder.ProgramFiles);
                        string[] guesses =
                        {
                            IOPath.Combine(pf, @"Autodesk\Inventor 2026\Bin\Public Assemblies\Autodesk.Inventor.Interop.dll"),
                            IOPath.Combine(pf, @"Autodesk\Inventor 2026\Bin\Autodesk.Inventor.Interop.dll")
                        };
                        foreach (var g in guesses)
                            if (IOFile.Exists(g)) return Assembly.LoadFrom(g);
                    }
                }
                catch { }
                return null;
            };

            try
            {
                MessageBox.Show("UCS Inspector: static ctor OK", "UcsInspectorperu",
                    MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            catch { }
            Log("static ctor reached");
        }

        private Inventor.Application _app;
        private ButtonDefinition _btn;
        private readonly string _clientId = typeof(StandardAddInServer).GUID.ToString("B");

        public StandardAddInServer()
        {
            Log("ctor StandardAddInServer()");
            try
            {
                MessageBox.Show("UCS Inspector: ctor", "UcsInspectorperu",
                    MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            catch { }
        }

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                Log("Activate() start. firstTime=" + firstTime);
                _app = addInSiteObject.Application;
                Log("Assembly.Location = " + typeof(StandardAddInServer).Assembly.Location);

                var defs = _app.CommandManager.ControlDefinitions;
                _btn = defs.AddButtonDefinition(
                    "UCS Inspector",
                    "UcsInspectorperu:UcsInspectorBtn",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    _clientId,
                    "Ver y editar offsets/ángulos de un UCS",
                    "Selecciona un UCS y aplica +, -, * o / a sus parámetros");

                _btn.OnExecute += (Context) =>
                {
                    try { new Ui.UcsForm(_app).Show(); }
                    catch (Exception ex) { Log("OnExecute error: " + ex); MessageBox.Show(ex.ToString()); }
                };

                TryAddToRibbon("Part", "id_TabTools", "UcsInspectorperu:PanelPart");
                TryAddToRibbon("Assembly", "id_TabTools", "UcsInspectorperu:PanelAsm");
                TryAddToRibbon("ZeroDoc", "id_TabToolsZeroDoc", "UcsInspectorperu:PanelZero");
                TryAddToRibbon("Drawing", "id_TabTools", "UcsInspectorperu:PanelDrw");

                Log("Activate() OK");
                MessageBox.Show("UCS Inspector: Activate() OK", "UcsInspectorperu",
                    MessageBoxButtons.OK, MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            catch (Exception ex)
            {
                Log("Activate() FAILED: " + ex);
                MessageBox.Show("UCS Inspector no pudo cargarse:\n\n" + ex, "UcsInspectorperu",
                    MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
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
                try { panel = tab.RibbonPanels["id_PanelProgramAddins"]; }
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
    }
}
