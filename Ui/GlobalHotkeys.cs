using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UcsInspectorperu.Ui // <-- ajusta si usas otro namespace
{
    internal static class GlobalHotkeys
    {
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int WM_HOTKEY = 0x0312;

        private sealed class MsgWnd : NativeWindow
        {
            private readonly Dictionary<int, Action> _map;
            public MsgWnd(IntPtr parent, Dictionary<int, Action> map)
            {
                _map = map;
                var cp = new CreateParams { Parent = parent };
                CreateHandle(cp);
            }
            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    int id = m.WParam.ToInt32();
                    Action a;
                    if (_map != null && _map.TryGetValue(id, out a))
                    {
                        try { a(); } catch { }
                    }
                }
                base.WndProc(ref m);
            }
        }

        private static MsgWnd _wnd;
        private static int _nextId;
        private static readonly Dictionary<int, Action> _actions = new Dictionary<int, Action>();

        public static void RegisterMany(IntPtr parent, IEnumerable<Tuple<uint, uint, Action>> combos)
        {
            UnregisterAll();
            _wnd = new MsgWnd(parent, _actions);
            _nextId = 1; _actions.Clear();

            foreach (var c in combos)
            {
                int id = _nextId++;
                _actions[id] = c.Item3;
                try { RegisterHotKey(_wnd.Handle, id, c.Item1, c.Item2); } catch { }
            }
        }

        public static void UnregisterAll()
        {
            try
            {
                if (_wnd != null)
                {
                    foreach (var id in new List<int>(_actions.Keys))
                    {
                        try { UnregisterHotKey(_wnd.Handle, id); } catch { }
                    }
                    _wnd.DestroyHandle();
                }
            }
            catch { }
            finally
            {
                _wnd = null;
                _actions.Clear();
                _nextId = 0;
            }
        }
    }
}
