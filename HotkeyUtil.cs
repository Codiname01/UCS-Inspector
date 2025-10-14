using System;
using System.Windows.Forms;

internal static class HotkeyUtil
{
    public struct Parsed
    {
        public Keys Key;
        public Keys ModMask;
        public uint FsModifiers;  // MOD_*
        public uint VirtualKey;   // VK_*
        public bool IsValid;
    }

    public static bool TryParse(string gesture, out Parsed p)
    {
        p = new Parsed(); if (string.IsNullOrWhiteSpace(gesture)) return false;
        string[] parts = gesture.Trim().Split('+');
        Keys key = Keys.None, mods = Keys.None; uint fs = 0, vk = 0;

        for (int i = 0; i < parts.Length; i++)
        {
            string up = parts[i].Trim().ToUpperInvariant();
            if (up == "ALT") { mods |= Keys.Alt; fs |= 0x0001; continue; }
            if (up == "CTRL" || up == "CONTROL") { mods |= Keys.Control; fs |= 0x0002; continue; }
            if (up == "SHIFT") { mods |= Keys.Shift; fs |= 0x0004; continue; }
            if (up == "WIN" || up == "WINDOWS") { fs |= 0x0008; continue; }

            if (up.Length == 1 && up[0] >= '0' && up[0] <= '9')
            { key = (Keys)((int)Keys.D0 + (up[0] - '0')); vk = (uint)key; continue; }

            if (up.StartsWith("NUM") && up.Length == 4 && char.IsDigit(up[3]))
            { int d = up[3] - '0'; key = (Keys)((int)Keys.NumPad0 + d); vk = (uint)key; continue; }

            if (up.StartsWith("F"))
            {
                int n; if (int.TryParse(up.Substring(1), out n) && n >= 1 && n <= 24)
                { key = Keys.F1 + (n - 1); vk = (uint)key; continue; }
            }

            if (up.Length == 1 && up[0] >= 'A' && up[0] <= 'Z')
            { key = (Keys)up[0]; vk = (uint)key; continue; }

            Keys temp;
            if (Enum.TryParse<Keys>(up, true, out temp))
            { key = temp; vk = (uint)key; continue; }

            return false;
        }

        if (key == Keys.None) return false;
        p.Key = key; p.ModMask = mods; p.FsModifiers = fs; p.VirtualKey = vk; p.IsValid = true; return true;
    }

    public static bool Matches(Keys keyData, string gesture)
    {
        Parsed p; if (!TryParse(gesture, out p)) return false;
        return ((keyData & Keys.KeyCode) == p.Key) && ((keyData & Keys.Modifiers) == p.ModMask);
    }
}
