#Requires -Version 5.1
<#
    RegDoktor 11  -  przeglądarka rejestru i centrum napraw dla Windows 11
    -------------------------------------------------------------------------
    * Przeglądarka rejestru (drzewo, wartości, edycja, wyszukiwanie, eksport/import .reg)
    * Wyszukiwanie i usuwanie osieroconych wpisów (zawsze z kopią .reg)
    * Analiza i kompaktowanie (defragmentacja) gałęzi rejestru
    * Menedżer autostartu
    * Kilkadziesiąt napraw w kolorowych kategoriach: system, sieć, dysk, RAM,
      wydajność, Eksplorator, urządzenia, bezpieczeństwo, rejestr

    Uruchomienie: dwuklik na "Uruchom RegDoktor.bat" (albo PPM na tym pliku > Uruchom za pomocą PowerShell).
    Wymaga uprawnień administratora - poprosi o nie automatycznie.
    Kopie zapasowe i raporty: Dokumenty\RegDoktor
#>
param([switch]$Relaunched)

$AppTitle = 'RegDoktor 11'
$AppVersion = '1.0.1'

# ============================================================== bootstrap
function Show-StartupError([string]$Message) {
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        [void][System.Windows.MessageBox]::Show($Message, $AppTitle, 'OK', 'Error')
    } catch { Write-Host $Message -ForegroundColor Red; Read-Host 'Naciśnij Enter' | Out-Null }
}

$principal = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$is64Ok = (-not [Environment]::Is64BitOperatingSystem) -or [Environment]::Is64BitProcess
$isWinPS = $PSVersionTable.PSVersion.Major -le 5
$isSTA = [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA

if (-not ($isAdmin -and $is64Ok -and $isWinPS -and $isSTA)) {
    if (-not $PSCommandPath) { Show-StartupError 'Uruchom RegDoktor jako plik skryptu (RegDoktor.ps1), a nie przez wklejenie kodu.'; exit 1 }
    if ($Relaunched) {
        Show-StartupError ("Nie udało się uruchomić RegDoktora w wymaganym trybie (administrator, 64-bit Windows PowerShell 5.1, STA).`n`nAdministrator: {0}`n64-bit: {1}`nWindows PowerShell: {2}`nSTA: {3}" -f $isAdmin, $is64Ok, $isWinPS, $isSTA)
        exit 1
    }
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        $psExe = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    } else {
        $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    }
    $argLine = '-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}" -Relaunched' -f $PSCommandPath
    try {
        if ($isAdmin) { Start-Process -FilePath $psExe -ArgumentList $argLine -ErrorAction Stop }
        else { Start-Process -FilePath $psExe -ArgumentList $argLine -Verb RunAs -ErrorAction Stop }
    } catch {
        Show-StartupError ("RegDoktor wymaga uprawnień administratora, aby naprawiać system i edytować rejestr.`n`n" + $_.Exception.Message)
    }
    exit
}

$ErrorActionPreference = 'Continue'

# ============================================================== C# helpers
$CSharpSource = @'
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32;

namespace RegDoktor
{
    public abstract class Bindable : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void Raise(string name)
        {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }
    }

    // ------------------------------------------------------------------ registry helpers
    public static class Reg
    {
        public static readonly string[] RootNames = new string[] {
            "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", "HKEY_USERS", "HKEY_CURRENT_CONFIG" };

        public static string ExpandRoot(string r)
        {
            if (r == null) return null;
            switch (r.Trim().TrimEnd(':').ToUpperInvariant())
            {
                case "HKCR": case "HKEY_CLASSES_ROOT": return "HKEY_CLASSES_ROOT";
                case "HKCU": case "HKEY_CURRENT_USER": return "HKEY_CURRENT_USER";
                case "HKLM": case "HKEY_LOCAL_MACHINE": return "HKEY_LOCAL_MACHINE";
                case "HKU": case "HKEY_USERS": return "HKEY_USERS";
                case "HKCC": case "HKEY_CURRENT_CONFIG": return "HKEY_CURRENT_CONFIG";
            }
            return null;
        }

        public static string NormalizePath(string path)
        {
            if (path == null) return null;
            string p = path.Trim().Trim('"').Replace('/', '\\').Trim('\\');
            if (p.Length == 0) return null;
            int idx = p.IndexOf('\\');
            string first = idx < 0 ? p : p.Substring(0, idx);
            string rest = idx < 0 ? "" : p.Substring(idx + 1);
            if (first.Equals("Computer", StringComparison.OrdinalIgnoreCase) ||
                first.Equals("Komputer", StringComparison.OrdinalIgnoreCase))
                return NormalizePath(rest);
            string root = ExpandRoot(first);
            if (root == null) return null;
            while (rest.IndexOf("\\\\", StringComparison.Ordinal) >= 0) rest = rest.Replace("\\\\", "\\");
            rest = rest.Trim('\\');
            return rest.Length == 0 ? root : root + "\\" + rest;
        }

        public static string ShortPath(string path)
        {
            string p = NormalizePath(path);
            if (p == null) return path;
            int idx = p.IndexOf('\\');
            string root = idx < 0 ? p : p.Substring(0, idx);
            string rest = idx < 0 ? "" : p.Substring(idx);
            string s;
            switch (root)
            {
                case "HKEY_CLASSES_ROOT": s = "HKCR"; break;
                case "HKEY_CURRENT_USER": s = "HKCU"; break;
                case "HKEY_LOCAL_MACHINE": s = "HKLM"; break;
                case "HKEY_USERS": s = "HKU"; break;
                default: s = "HKCC"; break;
            }
            return s + rest;
        }

        static RegistryHive HiveOf(string root)
        {
            switch (root)
            {
                case "HKEY_CLASSES_ROOT": return RegistryHive.ClassesRoot;
                case "HKEY_CURRENT_USER": return RegistryHive.CurrentUser;
                case "HKEY_LOCAL_MACHINE": return RegistryHive.LocalMachine;
                case "HKEY_USERS": return RegistryHive.Users;
                default: return RegistryHive.CurrentConfig;
            }
        }

        public static RegistryKey OpenKey(string path, bool writable)
        {
            string p = NormalizePath(path);
            if (p == null) return null;
            int idx = p.IndexOf('\\');
            string root = idx < 0 ? p : p.Substring(0, idx);
            string sub = idx < 0 ? "" : p.Substring(idx + 1);
            RegistryKey baseKey = RegistryKey.OpenBaseKey(HiveOf(root), RegistryView.Registry64);
            if (sub.Length == 0) return baseKey;
            try { return baseKey.OpenSubKey(sub, writable); }
            finally { baseKey.Close(); }
        }

        public static RegistryKey TryOpen(string path, bool writable)
        {
            try { return OpenKey(path, writable); } catch { return null; }
        }

        public static bool KeyExists(string path)
        {
            RegistryKey k = TryOpen(path, false);
            if (k == null) return false;
            k.Close();
            return true;
        }

        public static string Parent(string path)
        {
            string p = NormalizePath(path);
            if (p == null) return null;
            int i = p.LastIndexOf('\\');
            return i < 0 ? null : p.Substring(0, i);
        }

        public static string Leaf(string path)
        {
            string p = NormalizePath(path);
            if (p == null) return null;
            int i = p.LastIndexOf('\\');
            return i < 0 ? p : p.Substring(i + 1);
        }

        public static void DeleteKeyTree(string path)
        {
            string parent = Parent(path);
            string leaf = Leaf(path);
            if (parent == null) throw new InvalidOperationException("Root key cannot be deleted.");
            RegistryKey pk = OpenKey(parent, true);
            if (pk == null) throw new InvalidOperationException("Parent key not found: " + parent);
            using (pk) { pk.DeleteSubKeyTree(leaf, false); }
        }

        public static void DeleteValue(string path, string name)
        {
            RegistryKey k = OpenKey(path, true);
            if (k == null) throw new InvalidOperationException("Key not found: " + path);
            using (k) { k.DeleteValue(name ?? "", false); }
        }

        public static string CreateSubKey(string path, string name)
        {
            RegistryKey k = OpenKey(path, true);
            if (k == null) throw new InvalidOperationException("Key not found: " + path);
            using (k)
            {
                using (RegistryKey c = k.CreateSubKey(name)) { }
            }
            return NormalizePath(path) + "\\" + name;
        }

        public static string KindName(RegistryValueKind kind)
        {
            switch (kind)
            {
                case RegistryValueKind.String: return "REG_SZ";
                case RegistryValueKind.ExpandString: return "REG_EXPAND_SZ";
                case RegistryValueKind.DWord: return "REG_DWORD";
                case RegistryValueKind.QWord: return "REG_QWORD";
                case RegistryValueKind.MultiString: return "REG_MULTI_SZ";
                case RegistryValueKind.Binary: return "REG_BINARY";
                case RegistryValueKind.None: return "REG_NONE";
                default: return "REG_UNKNOWN";
            }
        }

        public static string Hex(byte[] b, int max)
        {
            if (b == null) return "";
            StringBuilder sb = new StringBuilder();
            int n = Math.Min(b.Length, max);
            for (int i = 0; i < n; i++)
            {
                if (i > 0) sb.Append(' ');
                sb.Append(b[i].ToString("x2"));
            }
            if (b.Length > max) sb.Append(" \u2026 (" + b.Length + " B)");
            return sb.ToString();
        }

        public static string Format(object v, RegistryValueKind kind)
        {
            if (v == null) return "";
            switch (kind)
            {
                case RegistryValueKind.String:
                case RegistryValueKind.ExpandString:
                    return v.ToString();
                case RegistryValueKind.DWord:
                    {
                        uint u = unchecked((uint)Convert.ToInt32(v));
                        return "0x" + u.ToString("x8") + " (" + u.ToString(CultureInfo.InvariantCulture) + ")";
                    }
                case RegistryValueKind.QWord:
                    {
                        ulong u = unchecked((ulong)Convert.ToInt64(v));
                        return "0x" + u.ToString("x16") + " (" + u.ToString(CultureInfo.InvariantCulture) + ")";
                    }
                case RegistryValueKind.MultiString:
                    {
                        string[] a = v as string[];
                        return a == null ? v.ToString() : string.Join(" | ", a);
                    }
                default:
                    {
                        byte[] b = v as byte[];
                        return b != null ? Hex(b, 256) : v.ToString();
                    }
            }
        }

        static RegValueRow MakeRow(RegistryKey k, string n)
        {
            RegValueRow r = new RegValueRow();
            r.Name = n;
            r.DisplayName = n.Length == 0 ? "(Domy\u015blna)" : n;
            try
            {
                r.Kind = k.GetValueKind(n);
                object v = k.GetValue(n, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                r.Data = Format(v, r.Kind);
            }
            catch (Exception ex)
            {
                r.Kind = RegistryValueKind.Unknown;
                r.Data = "(b\u0142\u0105d odczytu: " + ex.Message + ")";
            }
            r.KindName = KindName(r.Kind);
            r.IsString = r.Kind == RegistryValueKind.String || r.Kind == RegistryValueKind.ExpandString || r.Kind == RegistryValueKind.MultiString;
            return r;
        }

        public static List<RegValueRow> GetValues(string path)
        {
            List<RegValueRow> list = new List<RegValueRow>();
            RegistryKey k = OpenKey(path, false);
            if (k == null) return list;
            using (k)
            {
                string[] names = k.GetValueNames();
                bool hasDefault = false;
                foreach (string n in names) if (n.Length == 0) hasDefault = true;
                if (!hasDefault)
                {
                    RegValueRow d = new RegValueRow();
                    d.Name = ""; d.DisplayName = "(Domy\u015blna)"; d.Kind = RegistryValueKind.String;
                    d.KindName = "REG_SZ"; d.Data = "(warto\u015b\u0107 nieustawiona)"; d.IsString = true; d.IsUnset = true;
                    list.Add(d);
                }
                Array.Sort(names, StringComparer.OrdinalIgnoreCase);
                foreach (string n in names) list.Add(MakeRow(k, n));
            }
            return list;
        }

        public static string ToEditText(string path, string name, out RegistryValueKind kind)
        {
            kind = RegistryValueKind.String;
            RegistryKey k = OpenKey(path, false);
            if (k == null) return "";
            using (k)
            {
                object v = k.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (v == null) return "";
                kind = k.GetValueKind(name);
                switch (kind)
                {
                    case RegistryValueKind.DWord:
                        return unchecked((uint)Convert.ToInt32(v)).ToString(CultureInfo.InvariantCulture);
                    case RegistryValueKind.QWord:
                        return unchecked((ulong)Convert.ToInt64(v)).ToString(CultureInfo.InvariantCulture);
                    case RegistryValueKind.MultiString:
                        {
                            string[] a = v as string[];
                            return a == null ? "" : string.Join("\r\n", a);
                        }
                    case RegistryValueKind.String:
                    case RegistryValueKind.ExpandString:
                        return v.ToString();
                    default:
                        {
                            byte[] b = v as byte[];
                            return b == null ? v.ToString() : Hex(b, int.MaxValue);
                        }
                }
            }
        }

        public static byte[] ParseHex(string s)
        {
            List<byte> res = new List<byte>();
            StringBuilder clean = new StringBuilder();
            foreach (char c in s ?? "")
            {
                if (Uri.IsHexDigit(c)) clean.Append(c);
                else if (c == ' ' || c == ',' || c == '\r' || c == '\n' || c == '\t' || c == '-') continue;
                else throw new FormatException("Niedozwolony znak w danych binarnych: " + c);
            }
            string h = clean.ToString();
            if (h.Length % 2 != 0) throw new FormatException("Nieparzysta liczba cyfr szesnastkowych.");
            for (int i = 0; i < h.Length; i += 2)
                res.Add(byte.Parse(h.Substring(i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
            return res.ToArray();
        }

        static ulong ParseNumber(string text, ulong max)
        {
            string t = (text ?? "").Trim().Replace(" ", "");
            ulong u;
            bool ok;
            if (t.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                ok = ulong.TryParse(t.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out u);
            else
                ok = ulong.TryParse(t, NumberStyles.None, CultureInfo.InvariantCulture, out u);
            if (!ok || u > max) throw new FormatException("Nieprawid\u0142owa liczba: " + text);
            return u;
        }

        public static object ParseValue(string text, RegistryValueKind kind)
        {
            switch (kind)
            {
                case RegistryValueKind.DWord: return unchecked((int)(uint)ParseNumber(text, uint.MaxValue));
                case RegistryValueKind.QWord: return unchecked((long)ParseNumber(text, ulong.MaxValue));
                case RegistryValueKind.MultiString:
                    {
                        string t = (text ?? "").Replace("\r\n", "\n");
                        List<string> parts = new List<string>(t.Split('\n'));
                        while (parts.Count > 0 && parts[parts.Count - 1].Length == 0) parts.RemoveAt(parts.Count - 1);
                        return parts.ToArray();
                    }
                case RegistryValueKind.Binary:
                case RegistryValueKind.None:
                case RegistryValueKind.Unknown:
                    return ParseHex(text);
                default: return text ?? "";
            }
        }

        public static void SetValue(string path, string name, RegistryValueKind kind, string text)
        {
            object v = ParseValue(text, kind);
            if (kind == RegistryValueKind.Unknown) kind = RegistryValueKind.Binary;
            RegistryKey k = OpenKey(path, true);
            if (k == null) throw new InvalidOperationException("Key not found: " + path);
            using (k) { k.SetValue(name ?? "", v, kind); }
        }

        public static bool ValueExists(string path, string name)
        {
            RegistryKey k = TryOpen(path, false);
            if (k == null) return false;
            using (k)
            {
                foreach (string n in k.GetValueNames())
                    if (string.Equals(n, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        // ---- .reg export of a single key's values (no subkeys) - fast backup before value edits/deletions
        static string Esc(string s)
        {
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        static string HexList(byte[] b)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < b.Length; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append(b[i].ToString("x2"));
            }
            return sb.ToString();
        }

        static byte[] Utf16z(string s)
        {
            return Encoding.Unicode.GetBytes((s ?? "") + "\0");
        }

        static string EncodeRegValue(RegistryValueKind kind, object v)
        {
            switch (kind)
            {
                case RegistryValueKind.String:
                    {
                        string s = v as string ?? "";
                        if (s.IndexOfAny(new char[] { '\r', '\n', '\0' }) >= 0) return "hex(1):" + HexList(Utf16z(s));
                        return "\"" + Esc(s) + "\"";
                    }
                case RegistryValueKind.ExpandString:
                    return "hex(2):" + HexList(Utf16z(v as string));
                case RegistryValueKind.DWord:
                    return "dword:" + unchecked((uint)Convert.ToInt32(v)).ToString("x8");
                case RegistryValueKind.QWord:
                    return "hex(b):" + HexList(BitConverter.GetBytes(Convert.ToInt64(v)));
                case RegistryValueKind.MultiString:
                    {
                        string[] a = v as string[] ?? new string[0];
                        List<byte> bytes = new List<byte>();
                        foreach (string s in a) bytes.AddRange(Utf16z(s));
                        bytes.AddRange(new byte[] { 0, 0 });
                        return "hex(7):" + HexList(bytes.ToArray());
                    }
                case RegistryValueKind.Binary:
                    return "hex:" + HexList(v as byte[] ?? new byte[0]);
                case RegistryValueKind.None:
                    return "hex(0):" + HexList(v as byte[] ?? new byte[0]);
                default:
                    return null;
            }
        }

        public static int ExportKeyValues(string path, string file)
        {
            string p = NormalizePath(path);
            StringBuilder sb = new StringBuilder();
            sb.Append("Windows Registry Editor Version 5.00\r\n\r\n[" + p + "]\r\n");
            int count = 0;
            RegistryKey k = OpenKey(p, false);
            if (k == null) throw new InvalidOperationException("Key not found: " + p);
            using (k)
            {
                foreach (string n in k.GetValueNames())
                {
                    RegistryValueKind kind;
                    object v;
                    try
                    {
                        kind = k.GetValueKind(n);
                        v = k.GetValue(n, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                    }
                    catch { continue; }
                    string enc = EncodeRegValue(kind, v);
                    if (enc == null) { sb.Append("; skipped value of unsupported type: " + n + "\r\n"); continue; }
                    sb.Append(n.Length == 0 ? "@" : "\"" + Esc(n) + "\"");
                    sb.Append("=");
                    sb.Append(enc);
                    sb.Append("\r\n");
                    count++;
                }
            }
            sb.Append("\r\n");
            System.IO.File.WriteAllText(file, sb.ToString(), Encoding.Unicode);
            return count;
        }

        public static string FormatSize(long bytes)
        {
            double b = bytes;
            string[] u = new string[] { "B", "KB", "MB", "GB", "TB" };
            int i = 0;
            while (Math.Abs(b) >= 1024 && i < u.Length - 1) { b /= 1024; i++; }
            return b.ToString(i == 0 ? "0" : "0.0", CultureInfo.CurrentCulture) + " " + u[i];
        }

        // For each child key of rootPath: read child\relSub value "valueName". Returns [childName, value].
        public static List<string[]> CollectChildValues(string rootPath, string relSub, string valueName)
        {
            List<string[]> res = new List<string[]>();
            RegistryKey root = TryOpen(rootPath, false);
            if (root == null) return res;
            using (root)
            {
                string[] names;
                try { names = root.GetSubKeyNames(); } catch { return res; }
                foreach (string n in names)
                {
                    try
                    {
                        string rel = string.IsNullOrEmpty(relSub) ? n : n + "\\" + relSub;
                        using (RegistryKey k = root.OpenSubKey(rel, false))
                        {
                            if (k == null) continue;
                            object v = k.GetValue(valueName ?? "", null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                            string s = v as string;
                            if (s != null && s.Length > 0) res.Add(new string[] { n, s });
                        }
                    }
                    catch { }
                }
            }
            return res;
        }

        // COM classes: [clsid, serverKind, path]
        public static List<string[]> CollectComServers(string rootPath)
        {
            List<string[]> res = new List<string[]>();
            RegistryKey root = TryOpen(rootPath, false);
            if (root == null) return res;
            string[] kinds = new string[] { "InprocServer32", "LocalServer32" };
            using (root)
            {
                string[] names;
                try { names = root.GetSubKeyNames(); } catch { return res; }
                foreach (string n in names)
                {
                    try
                    {
                        using (RegistryKey k = root.OpenSubKey(n, false))
                        {
                            if (k == null) continue;
                            foreach (string kind in kinds)
                            {
                                using (RegistryKey sk = k.OpenSubKey(kind, false))
                                {
                                    if (sk == null) continue;
                                    object v = sk.GetValue("", null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                                    string s = v as string;
                                    if (s != null && s.Length > 0) { res.Add(new string[] { n, kind, s }); break; }
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            return res;
        }

        // Keys with no values and no subkeys, up to 'depth' levels below rootPath.
        public static List<string> FindEmptyKeys(string rootPath, int depth, string[] excludeTop)
        {
            List<string> res = new List<string>();
            RegistryKey root = TryOpen(rootPath, false);
            if (root == null) return res;
            using (root) { FindEmpty(root, NormalizePath(rootPath), depth, excludeTop, res); }
            return res;
        }

        static void FindEmpty(RegistryKey k, string path, int depth, string[] exclude, List<string> res)
        {
            if (depth <= 0) return;
            string[] names;
            try { names = k.GetSubKeyNames(); } catch { return; }
            foreach (string n in names)
            {
                if (exclude != null)
                {
                    bool skip = false;
                    foreach (string e in exclude) if (string.Equals(e, n, StringComparison.OrdinalIgnoreCase)) { skip = true; break; }
                    if (skip) continue;
                }
                try
                {
                    using (RegistryKey c = k.OpenSubKey(n, false))
                    {
                        if (c == null) continue;
                        if (c.SubKeyCount == 0 && c.ValueCount == 0) res.Add(path + "\\" + n);
                        else FindEmpty(c, path + "\\" + n, depth - 1, null, res);
                    }
                }
                catch { }
            }
        }
    }

    // ------------------------------------------------------------------ data rows
    public class RegValueRow
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string KindName { get; set; }
        public RegistryValueKind Kind { get; set; }
        public string Data { get; set; }
        public bool IsString { get; set; }
        public bool IsUnset { get; set; }
        public string Icon { get { return IsString ? "ab" : "01"; } }
    }

    public class RegNode : Bindable
    {
        private bool _expanded;
        private bool _selected;
        private ObservableCollection<RegNode> _children;

        public string Name { get; set; }
        public string FullPath { get; set; }
        public bool IsDummy { get; set; }
        public bool IsLoaded { get; private set; }
        public RegNode ParentNode { get; set; }

        public ObservableCollection<RegNode> Children
        {
            get { return _children; }
            private set { _children = value; Raise("Children"); }
        }

        public RegNode()
        {
            _children = new ObservableCollection<RegNode>();
        }

        public RegNode(string name, string fullPath, RegNode parent, bool addDummy)
        {
            Name = name;
            FullPath = fullPath;
            ParentNode = parent;
            _children = new ObservableCollection<RegNode>();
            if (addDummy)
            {
                RegNode d = new RegNode();
                d.Name = "\u2026";
                d.IsDummy = true;
                _children.Add(d);
            }
        }

        public bool IsExpanded
        {
            get { return _expanded; }
            set
            {
                if (_expanded == value) return;
                _expanded = value;
                if (value && !IsLoaded) Reload();
                Raise("IsExpanded");
            }
        }

        public bool IsSelected
        {
            get { return _selected; }
            set { if (_selected == value) return; _selected = value; Raise("IsSelected"); }
        }

        public void EnsureLoaded()
        {
            if (!IsLoaded) Reload();
        }

        public void Reload()
        {
            IsLoaded = true;
            List<RegNode> items = new List<RegNode>();
            RegistryKey k = Reg.TryOpen(FullPath, false);
            if (k != null)
            {
                using (k)
                {
                    string[] names;
                    try { names = k.GetSubKeyNames(); } catch { names = new string[0]; }
                    Array.Sort(names, StringComparer.OrdinalIgnoreCase);
                    bool probe = names.Length <= 3000;
                    foreach (string n in names)
                    {
                        bool has = true;
                        if (probe)
                        {
                            try
                            {
                                using (RegistryKey c = k.OpenSubKey(n, false))
                                {
                                    has = c != null && c.SubKeyCount > 0;
                                }
                            }
                            catch { has = true; }
                        }
                        items.Add(new RegNode(n, FullPath + "\\" + n, this, has));
                    }
                }
            }
            Children = new ObservableCollection<RegNode>(items);
        }

        public RegNode FindChild(string name)
        {
            EnsureLoaded();
            foreach (RegNode c in Children)
                if (!c.IsDummy && string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)) return c;
            return null;
        }

        public static ObservableCollection<RegNode> CreateRoots()
        {
            ObservableCollection<RegNode> roots = new ObservableCollection<RegNode>();
            foreach (string r in Reg.RootNames) roots.Add(new RegNode(r, r, null, true));
            return roots;
        }
    }

    public class CleanItem : Bindable
    {
        private bool _selected;
        public bool Selected { get { return _selected; } set { if (_selected != value) { _selected = value; Raise("Selected"); } } }
        public string Category { get; set; }
        public string Risk { get; set; }
        public int RiskLevel { get; set; }
        public string KeyPath { get; set; }
        public string ValueName { get; set; }
        public bool IsWholeKey { get; set; }
        public string Target { get; set; }
        public string Detail { get; set; }
        public string What { get { return IsWholeKey ? "(ca\u0142y klucz)" : (string.IsNullOrEmpty(ValueName) ? "(Domy\u015blna)" : ValueName); } }
    }

    public class StartupItem
    {
        public string Name { get; set; }
        public string Command { get; set; }
        public string Location { get; set; }
        public string Scope { get; set; }
        public string RunKeyPath { get; set; }      // registry key holding the entry (or null for folder items)
        public string FilePath { get; set; }        // startup-folder file (or null)
        public string ApprovedKeyPath { get; set; } // StartupApproved key
        public bool Enabled { get; set; }
        public bool Missing { get; set; }
        public string Status { get { return (Enabled ? "W\u0142\u0105czony" : "Wy\u0142\u0105czony") + (Missing ? "  \u2022  brak pliku!" : ""); } }
    }

    public class HiveInfo : Bindable
    {
        private bool _selected;
        private string _status;
        public bool Selected { get { return _selected; } set { if (_selected != value) { _selected = value; Raise("Selected"); } } }
        public string Name { get; set; }
        public string Root { get; set; }     // HKLM / HKU
        public string SubKey { get; set; }   // SOFTWARE, S-1-5-21-..., ...
        public string FilePath { get; set; }
        public long Size { get; set; }
        public long Compacted { get; set; }
        public string Status { get { return _status; } set { _status = value; Raise("Status"); } }
        public string SizeText { get { return Size > 0 ? Reg.FormatSize(Size) : "-"; } }
        public string CompactedText { get { return Compacted > 0 ? Reg.FormatSize(Compacted) : "-"; } }
        public double SavingPercent { get { return (Size > 0 && Compacted > 0) ? Math.Round(100.0 * (Size - Compacted) / Size, 1) : 0; } }
        public string SavingText { get { return (Size > 0 && Compacted > 0) ? (Reg.FormatSize(Size - Compacted) + "  (" + SavingPercent.ToString("0.0", CultureInfo.CurrentCulture) + "%)") : "-"; } }
    }

    public class LogLine
    {
        public string Time { get; set; }
        public string Text { get; set; }
        public string Kind { get; set; }
        public string Display { get { return "[" + Time + "]  " + Text; } }
        public LogLine(string text, string kind)
        {
            Time = DateTime.Now.ToString("HH:mm:ss");
            Text = text;
            Kind = kind;
        }
    }

    // ------------------------------------------------------------------ registry search (background thread)
    public class SearchHit
    {
        public string Path { get; set; }
        public string Where { get; set; }
        public string ValueName { get; set; }
        public string Data { get; set; }
    }

    public class RegSearch
    {
        private volatile bool _cancel;
        private volatile bool _running;
        private long _scanned;
        private int _found;
        private volatile string _current = "";
        private string _text;
        private bool _keys, _names, _data;
        public ConcurrentQueue<SearchHit> Results = new ConcurrentQueue<SearchHit>();
        public int MaxResults = 5000;

        public bool IsRunning { get { return _running; } }
        public long Scanned { get { return Interlocked.Read(ref _scanned); } }
        public int Found { get { return _found; } }
        public string Current { get { return _current; } }

        public void Cancel() { _cancel = true; }

        public void Start(string[] roots, string text, bool keys, bool names, bool data)
        {
            if (_running) return;
            _cancel = false; _running = true; _scanned = 0; _found = 0;
            _text = text; _keys = keys; _names = names; _data = data;
            SearchHit dummy;
            while (Results.TryDequeue(out dummy)) { }
            string[] rootsCopy = (string[])roots.Clone();
            Thread t = new Thread(delegate ()
            {
                try
                {
                    foreach (string r in rootsCopy)
                    {
                        if (_cancel) break;
                        string p = Reg.NormalizePath(r);
                        RegistryKey k = Reg.TryOpen(p, false);
                        if (k == null) continue;
                        using (k) { Walk(k, p); }
                    }
                }
                catch { }
                finally { _running = false; }
            }, 16 * 1024 * 1024);
            t.IsBackground = true;
            t.Start();
        }

        private void Add(string path, string where, string name, string data)
        {
            SearchHit h = new SearchHit();
            h.Path = path; h.Where = where; h.ValueName = name; h.Data = data;
            Results.Enqueue(h);
            Interlocked.Increment(ref _found);
        }

        private static string Leaf(string path)
        {
            int i = path.LastIndexOf('\\');
            return i < 0 ? path : path.Substring(i + 1);
        }

        private static string DataText(RegistryKey k, string n)
        {
            try
            {
                object v = k.GetValue(n, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (v == null) return null;
                string s = v as string;
                if (s != null) return s;
                string[] a = v as string[];
                if (a != null) return string.Join(" | ", a);
                if (v is int || v is long) return Reg.Format(v, v is int ? RegistryValueKind.DWord : RegistryValueKind.QWord);
                return null;
            }
            catch { return null; }
        }

        private void Walk(RegistryKey key, string path)
        {
            if (_cancel || _found >= MaxResults) return;
            Interlocked.Increment(ref _scanned);
            _current = path;
            if (_keys && Leaf(path).IndexOf(_text, StringComparison.OrdinalIgnoreCase) >= 0)
                Add(path, "Klucz", null, null);
            if (_names || _data)
            {
                string[] vn;
                try { vn = key.GetValueNames(); } catch { vn = new string[0]; }
                foreach (string n in vn)
                {
                    if (_names && n.Length > 0 && n.IndexOf(_text, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        Add(path, "Nazwa warto\u015bci", n, DataText(key, n));
                        continue;
                    }
                    if (_data)
                    {
                        string d = DataText(key, n);
                        if (d != null && d.IndexOf(_text, StringComparison.OrdinalIgnoreCase) >= 0)
                            Add(path, "Dane", n.Length == 0 ? "(Domy\u015blna)" : n, d);
                    }
                }
            }
            string[] subs;
            try { subs = key.GetSubKeyNames(); } catch { return; }
            Array.Sort(subs, StringComparer.OrdinalIgnoreCase);
            foreach (string s in subs)
            {
                if (_cancel || _found >= MaxResults) return;
                string childPath = path + "\\" + s;
                // CurrentControlSet is a symbolic link to ControlSet00x - avoid scanning it twice
                if (string.Equals(childPath, "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet", StringComparison.OrdinalIgnoreCase)) continue;
                RegistryKey c = null;
                try { c = key.OpenSubKey(s, false); } catch { c = null; }
                if (c == null) continue;
                try { Walk(c, childPath); }
                finally { c.Close(); }
            }
        }
    }

    // ------------------------------------------------------------------ native helpers
    public static class Native
    {
        [StructLayout(LayoutKind.Sequential)]
        struct LUID { public uint Low; public int High; }

        [StructLayout(LayoutKind.Sequential)]
        struct TOKEN_PRIVILEGES { public uint Count; public LUID Luid; public uint Attr; }

        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool LookupPrivilegeValue(string system, string name, out LUID luid);

        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool AdjustTokenPrivileges(IntPtr token, bool disableAll, ref TOKEN_PRIVILEGES state, uint len, IntPtr prev, IntPtr retLen);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool CloseHandle(IntPtr h);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        static extern int RegCreateKeyEx(IntPtr hKey, string subKey, int reserved, string cls, int options, int sam, IntPtr security, out IntPtr result, out int disposition);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        static extern int RegOpenKeyEx(IntPtr hKey, string subKey, int options, int sam, out IntPtr result);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        static extern int RegSaveKeyEx(IntPtr hKey, string file, IntPtr security, int flags);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        static extern int RegReplaceKey(IntPtr hKey, string subKey, string newFile, string oldFile);

        [DllImport("advapi32.dll")]
        static extern int RegCloseKey(IntPtr hKey);

        [DllImport("ntdll.dll")]
        static extern int NtSetSystemInformation(int infoClass, ref int info, int length);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern uint QueryDosDevice(string device, StringBuilder target, int max);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int cmd);

        [DllImport("shell32.dll")]
        static extern void SHChangeNotify(int eventId, uint flags, IntPtr item1, IntPtr item2);

        [DllImport("kernel32.dll")]
        static extern uint GetOEMCP();

        [DllImport("dwmapi.dll")]
        static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        // Dark title bar on Windows 10 20H1+ / Windows 11 (DWMWA_USE_IMMERSIVE_DARK_MODE = 20)
        public static void DarkTitleBar(IntPtr hwnd)
        {
            try
            {
                int on = 1;
                if (DwmSetWindowAttribute(hwnd, 20, ref on, 4) != 0) DwmSetWindowAttribute(hwnd, 19, ref on, 4);
            }
            catch { }
        }

        public static int OemCodePage()
        {
            return (int)GetOEMCP();
        }

        public static readonly IntPtr HKLM = new IntPtr(unchecked((int)0x80000002));
        public static readonly IntPtr HKU = new IntPtr(unchecked((int)0x80000003));

        public static bool EnablePrivilege(string name)
        {
            IntPtr tok;
            if (!OpenProcessToken(GetCurrentProcess(), 0x0020 | 0x0008, out tok)) return false;
            try
            {
                LUID l;
                if (!LookupPrivilegeValue(null, name, out l)) return false;
                TOKEN_PRIVILEGES tp = new TOKEN_PRIVILEGES();
                tp.Count = 1; tp.Luid = l; tp.Attr = 0x2;
                if (!AdjustTokenPrivileges(tok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero)) return false;
                return Marshal.GetLastWin32Error() == 0;
            }
            finally { CloseHandle(tok); }
        }

        static IntPtr RootHandle(string root)
        {
            string r = Reg.ExpandRoot(root);
            if (r == "HKEY_LOCAL_MACHINE") return HKLM;
            if (r == "HKEY_USERS") return HKU;
            throw new ArgumentException("Only HKLM or HKU hives are supported: " + root);
        }

        // Saves a hive (root + hive subkey) to a new, compacted file. Returns Win32 error code (0 = OK).
        public static int SaveHive(string root, string subKey, string file)
        {
            EnablePrivilege("SeBackupPrivilege");
            IntPtr h;
            int rc = RegOpenKeyEx(RootHandle(root), subKey, 0, 0x20019, out h);
            if (rc != 0)
            {
                int disp;
                rc = RegCreateKeyEx(RootHandle(root), subKey, 0, null, 4, 0x20019, IntPtr.Zero, out h, out disp);
                if (rc != 0) return rc;
            }
            try { return RegSaveKeyEx(h, file, IntPtr.Zero, 2); }
            finally { RegCloseKey(h); }
        }

        // Schedules a hive file replacement (takes effect after restart). Returns Win32 error code.
        public static int ReplaceHive(string root, string subKey, string newFile, string oldFile)
        {
            EnablePrivilege("SeBackupPrivilege");
            EnablePrivilege("SeRestorePrivilege");
            return RegReplaceKey(RootHandle(root), subKey, newFile, oldFile);
        }

        // 4 = MemoryPurgeStandbyList, 2 = MemoryEmptyWorkingSets. Returns NTSTATUS (0 = OK).
        public static int MemoryCommand(int command)
        {
            EnablePrivilege("SeProfileSingleProcessPrivilege");
            EnablePrivilege("SeIncreaseQuotaPrivilege");
            int cmd = command;
            return NtSetSystemInformation(80, ref cmd, 4);
        }

        public static string DeviceToDos(string devicePath)
        {
            if (string.IsNullOrEmpty(devicePath)) return devicePath;
            foreach (string drive in Environment.GetLogicalDrives())
            {
                string d = drive.Substring(0, 2);
                StringBuilder sb = new StringBuilder(1024);
                if (QueryDosDevice(d, sb, sb.Capacity) == 0) continue;
                string dev = sb.ToString();
                if (dev.Length > 0 && devicePath.StartsWith(dev + "\\", StringComparison.OrdinalIgnoreCase))
                    return d + devicePath.Substring(dev.Length);
            }
            return devicePath;
        }

        public static void HideConsole()
        {
            IntPtr h = GetConsoleWindow();
            if (h != IntPtr.Zero) ShowWindow(h, 0);
        }

        public static void RefreshShellAssociations()
        {
            SHChangeNotify(0x08000000, 0, IntPtr.Zero, IntPtr.Zero);
        }
    }
}
'@

try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml -ErrorAction Stop
    if (-not ('RegDoktor.Reg' -as [type])) { Add-Type -TypeDefinition $CSharpSource -Language CSharp -ErrorAction Stop }
} catch {
    Show-StartupError ("Nie udało się załadować komponentów programu:`n`n" + $_.Exception.Message)
    exit 1
}
try { [RegDoktor.Native]::HideConsole() } catch { }

# ============================================================== paths & shared state
$DocsDir = [Environment]::GetFolderPath('MyDocuments')
if (-not $DocsDir) { $DocsDir = $env:USERPROFILE }
$AppDataRoot = Join-Path $DocsDir 'RegDoktor'
$BackupRoot = Join-Path $AppDataRoot 'Kopie zapasowe'
$ReportRoot = Join-Path $AppDataRoot 'Raporty'
foreach ($d in @($AppDataRoot, $BackupRoot, $ReportRoot)) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }
$LogFile = Join-Path $AppDataRoot ('Dziennik_' + (Get-Date -Format 'yyyy-MM') + '.log')

try { $OemEncoding = [System.Text.Encoding]::GetEncoding([RegDoktor.Native]::OemCodePage()) } catch { $OemEncoding = [System.Text.Encoding]::GetEncoding(852) }

$Sync = [hashtable]::Synchronized(@{
    Queue  = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    Cancel = $false
    Proc   = $null
})

$State = @{
    Job = $null; CancelAt = $null
    RegPath = $null; PendingChain = $null
    SearchWasRunning = $false; SearchStart = Get-Date
    LogDirty = $false
    LogBuffer = New-Object 'System.Collections.Generic.List[string]'
    Tiles = New-Object 'System.Collections.Generic.List[object]'
    Sections = New-Object 'System.Collections.Generic.List[object]'
    BusyControls = @()
}

# ============================================================== shared helpers (UI + background)
$SharedCode = {
    # ---- helpers shared by the UI thread and background tasks ----
    $script:FixedDriveCache = @{}

    function Format-Size([double]$Bytes) {
        return [RegDoktor.Reg]::FormatSize([long]$Bytes)
    }

    function Test-FixedDrive([string]$Letter) {
        $l = $Letter.ToUpperInvariant()
        if (-not $script:FixedDriveCache.ContainsKey($l)) {
            $ok = $false
            try {
                $d = New-Object System.IO.DriveInfo ($l + ':\')
                $ok = ($d.DriveType -eq [System.IO.DriveType]::Fixed) -and $d.IsReady
            } catch { $ok = $false }
            $script:FixedDriveCache[$l] = $ok
        }
        return $script:FixedDriveCache[$l]
    }

    # Extracts an absolute file path from a command line / registry value.
    # Returns $null when the path cannot be determined reliably (relative names, network paths,
    # removable drives, unexpanded variables) - such entries are never reported as orphans.
    function Get-CmdPath([string]$Text) {
        try {
            if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
            $s = [Environment]::ExpandEnvironmentVariables($Text).Trim()
            if ($s.StartsWith('@')) { $s = $s.Substring(1) }
            if ($s.StartsWith('\??\')) { $s = $s.Substring(4) }
            if ($s -match '^(?i)\\SystemRoot\\') { $s = $env:SystemRoot + $s.Substring(11) }
            if ($s -match '^(?i)(System32|SysWOW64)\\') { $s = [System.IO.Path]::Combine($env:SystemRoot, $s) }
            $rest = ''
            if ($s.StartsWith('"')) {
                $e = $s.IndexOf('"', 1)
                if ($e -lt 0) { $p = $s.Substring(1) } else { $p = $s.Substring(1, $e - 1); $rest = $s.Substring($e + 1) }
            } else {
                $m = [regex]::Match($s, '^(.+?\.(?:exe|dll|ocx|cpl|com|bat|cmd|msc|scr|sys|vbs|vbe|js|jse|wsf|ps1|lnk|ax|drv|tlb|olb|hta|msi|jar))(?=$|[\s,"/])', 'IgnoreCase')
                if ($m.Success) { $p = $m.Groups[1].Value; $rest = $s.Substring($m.Length) } else { $p = $s }
            }
            $p = $p.Trim()
            if ($p.Length -lt 4) { return $null }
            $leaf = [System.IO.Path]::GetFileName($p)
            if ($leaf -match '^(?i)(rundll32|regsvr32)(\.exe)?$') {
                $r = ($rest.Trim() -replace '^(?:/\w+\s+)+', '').Trim()
                if ($r) { return (Get-CmdPath $r) }
                return $null
            }
            if ($p -notmatch '^[A-Za-z]:\\') { return $null }
            if (-not (Test-FixedDrive $p.Substring(0, 1))) { return $null }
            if ($p.IndexOfAny([System.IO.Path]::GetInvalidPathChars()) -ge 0) { return $null }
            if ($p.Contains('%') -or $p.Contains('*') -or $p.Contains('?')) { return $null }
            return $p
        } catch { return $null }
    }

    # File/folder existence check tolerant to WOW64 redirection (System32/SysWOW64, Program Files/(x86)).
    function Test-PathExistsSmart([string]$Path) {
        if ([System.IO.File]::Exists($Path) -or [System.IO.Directory]::Exists($Path)) { return $true }
        $alts = New-Object System.Collections.Generic.List[string]
        $sys = [Environment]::GetFolderPath('System')
        $wow = [Environment]::GetFolderPath('SystemX86')
        $pf = $env:ProgramW6432
        if (-not $pf) { $pf = $env:ProgramFiles }
        $pf86 = ${env:ProgramFiles(x86)}
        $ic = [System.StringComparison]::OrdinalIgnoreCase
        if ($wow -and $Path.StartsWith($sys + '\', $ic)) { $alts.Add($wow + $Path.Substring($sys.Length)) }
        if ($pf86) {
            if ($Path.StartsWith($pf86 + '\', $ic)) { $alts.Add($pf + $Path.Substring($pf86.Length)) }
            elseif ($Path.StartsWith($pf + '\', $ic)) { $alts.Add($pf86 + $Path.Substring($pf.Length)) }
        }
        foreach ($a in $alts) { if ([System.IO.File]::Exists($a) -or [System.IO.Directory]::Exists($a)) { return $true } }
        return $false
    }

    # Returns the missing path when the command points to a file that does not exist, otherwise $null.
    function Get-MissingTarget([string]$Command) {
        $p = Get-CmdPath $Command
        if (-not $p) { return $null }
        if (Test-PathExistsSmart $p) { return $null }
        return $p
    }

    function ConvertTo-PackedGuid([string]$Guid) {
        $h = $Guid.Trim().Trim('{', '}') -replace '-', ''
        if ($h.Length -ne 32 -or $h -notmatch '^[0-9A-Fa-f]{32}$') { return $null }
        $a = -join ($h.Substring(0, 8).ToCharArray()[7..0])
        $b = -join ($h.Substring(8, 4).ToCharArray()[3..0])
        $c = -join ($h.Substring(12, 4).ToCharArray()[3..0])
        $d = ''
        for ($i = 16; $i -lt 32; $i += 2) { $d += $h.Substring($i + 1, 1) + $h.Substring($i, 1) }
        return ($a + $b + $c + $d).ToUpperInvariant()
    }

    function Get-SessionExplorer {
        $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
        return @(Get-Process -Name explorer -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sid })
    }

    # Opens a file/URL through the (non-elevated) shell, e.g. HTML reports in the default browser.
    function Open-Unelevated([string]$Target) {
        try { Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList ('"' + $Target + '"') }
        catch { Start-Process -FilePath $Target }
    }
}
. $SharedCode

$WorkerPrelude = $SharedCode.ToString() + "`n" + {
    # ---- runs inside a background runspace; talks to the UI only through $Sync.Queue ----
    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'
    $WarningPreference = 'SilentlyContinue'

    function Log([string]$Text, [string]$Kind = 'info') { $Sync.Queue.Enqueue(@{ T = 'log'; Text = $Text; Kind = $Kind }) }
    function Head([string]$Text) { Log $Text 'head' }
    function Ok([string]$Text) { Log ('[OK] ' + $Text) 'ok' }
    function Warn([string]$Text) { Log ('[!] ' + $Text) 'warn' }
    function Err([string]$Text) { Log ('[X] ' + $Text) 'err' }
    function Progress([double]$Value = -1, [string]$Text = $null) { $Sync.Queue.Enqueue(@{ T = 'progress'; Value = $Value; Text = $Text }) }
    function Test-Cancel { return [bool]$Sync.Cancel }

    # Runs a console program, streams its output to the log (OEM code page, or UTF-16 for sfc).
    function Invoke-Cmd {
        param([string]$File, [string]$Arguments = '', [switch]$Unicode, [switch]$Quiet, [string]$WorkDir = $null)
        if (-not $Quiet) { Log ('> ' + [System.IO.Path]::GetFileName($File) + ' ' + $Arguments) 'cmd' }
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $File
        $psi.Arguments = $Arguments
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        if ($WorkDir) { $psi.WorkingDirectory = $WorkDir }
        if ($Unicode) { $enc = [System.Text.Encoding]::Unicode } else { $enc = $OemEncoding }
        $psi.StandardOutputEncoding = $enc
        $psi.StandardErrorEncoding = $enc
        try { $p = [System.Diagnostics.Process]::Start($psi) }
        catch { Err ('Nie można uruchomić ' + $File + ': ' + $_.Exception.Message); return -1 }
        $Sync.Proc = $p
        $errTask = $p.StandardError.ReadToEndAsync()
        $buffer = New-Object System.Collections.Generic.List[string]
        while ($true) {
            $line = $p.StandardOutput.ReadLine()
            if ($null -eq $line) { break }
            $t = ($line -replace "`0", '').Trim()
            if ($t.Length -eq 0) { continue }
            $pm = [regex]::Match($t, '(\d{1,3}(?:[.,]\d+)?)\s?%')
            if ($t.Length -lt 140 -and $pm.Success -and
                ($t.StartsWith('[') -or $t -match '(?i)(complete|uko.cz|zako.cz|weryfik|verif|post.p|progress|etap|stage|eta:)')) {
                $pct = 0.0
                [void][double]::TryParse(($pm.Groups[1].Value -replace ',', '.'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$pct)
                Progress $pct $t
                continue
            }
            if ($Quiet) { $buffer.Add($t) } else { Log $t 'out' }
        }
        $p.WaitForExit()
        $errText = ''
        try { $errText = $errTask.Result } catch { }
        $code = $p.ExitCode
        $Sync.Proc = $null
        $p.Dispose()
        if ($Quiet -and $code -ne 0) { foreach ($b in $buffer) { Log $b 'out' } }
        if ($errText) {
            foreach ($l in ($errText -split "`r?`n")) {
                $x = ($l -replace "`0", '').Trim()
                if ($x) { if ($code -eq 0) { Log $x 'out' } else { Log $x 'warn' } }
            }
        }
        return $code
    }

    function Stop-ServiceSafe([string]$Name) {
        $s = Get-Service -Name $Name -ErrorAction SilentlyContinue
        if (-not $s) { return }
        if ($s.Status -eq 'Stopped') { return }
        try { Stop-Service -Name $Name -Force -ErrorAction Stop; Log ('Zatrzymano usługę ' + $Name) }
        catch { Warn ('Nie udało się zatrzymać ' + $Name + ': ' + $_.Exception.Message) }
    }

    function Start-ServiceSafe([string]$Name) {
        $s = Get-Service -Name $Name -ErrorAction SilentlyContinue
        if (-not $s) { Warn ('Brak usługi ' + $Name); return }
        try {
            if ($s.StartType -eq 'Disabled') { Set-Service -Name $Name -StartupType Manual -ErrorAction Stop }
            if ($s.Status -ne 'Running') { Start-Service -Name $Name -ErrorAction Stop }
            Log ('Uruchomiono usługę ' + $Name)
        } catch { Warn ('Nie udało się uruchomić ' + $Name + ': ' + $_.Exception.Message) }
    }

    function Get-Stamp { return (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss') }

    function Get-ChangeBackupDir {
        if (-not $script:ChangeDir) {
            $script:ChangeDir = Join-Path $BackupRoot ('Zmiany_' + (Get-Stamp))
            New-Item -ItemType Directory -Path $script:ChangeDir -Force | Out-Null
        }
        return $script:ChangeDir
    }

    # Exports a registry key (with subkeys) to a .reg file in the change-backup folder.
    function Backup-RegKey([string]$Path, [string]$Label) {
        if (-not [RegDoktor.Reg]::KeyExists($Path)) { return }
        $file = Join-Path (Get-ChangeBackupDir) ($Label + '.reg')
        $rc = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\reg.exe') ('export "{0}" "{1}" /y' -f [RegDoktor.Reg]::ShortPath($Path), $file) -Quiet
        if ($rc -eq 0) { Log ('Kopia klucza: ' + $file) 'out' } else { Warn ('Nie udało się wykonać kopii ' + $Path) }
    }

    function Export-RegTree([string]$Path, [string]$File) {
        return (Invoke-Cmd (Join-Path $env:SystemRoot 'System32\reg.exe') ('export "{0}" "{1}" /y' -f [RegDoktor.Reg]::ShortPath($Path), $File) -Quiet)
    }

    function Set-Reg([string]$Path, [string]$Name, $Value, [string]$Type = 'DWord') {
        if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
        New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
    }

    function Get-RegValue([string]$Path, [string]$Name) {
        try { return (Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop).$Name } catch { return $null }
    }

    function New-RestorePointSafe([string]$Description) {
        Log 'Tworzenie punktu przywracania systemu (może potrwać do minuty)…'
        Progress -1 'Tworzenie punktu przywracania…'
        $rpKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
        $old = Get-RegValue $rpKey 'SystemRestorePointCreationFrequency'
        try {
            Enable-ComputerRestore -Drive ($env:SystemDrive + '\') -ErrorAction SilentlyContinue
            Set-Reg $rpKey 'SystemRestorePointCreationFrequency' 0 'DWord'
            Checkpoint-Computer -Description $Description -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
            Ok 'Punkt przywracania utworzony.'
            return $true
        } catch {
            Warn ('Nie udało się utworzyć punktu przywracania: ' + $_.Exception.Message)
            return $false
        } finally {
            try {
                if ($null -eq $old) { Remove-ItemProperty -LiteralPath $rpKey -Name 'SystemRestorePointCreationFrequency' -ErrorAction SilentlyContinue }
                else { Set-Reg $rpKey 'SystemRestorePointCreationFrequency' $old 'DWord' }
            } catch { }
        }
    }

    # Enumerates files and folders below a folder without following junctions/symlinks.
    function Get-TreeSafe([string]$Path) {
        $files = New-Object System.Collections.Generic.List[System.IO.FileInfo]
        $dirs = New-Object System.Collections.Generic.List[System.IO.DirectoryInfo]
        if ([System.IO.Directory]::Exists($Path)) {
            $stack = New-Object System.Collections.Stack
            $stack.Push((New-Object System.IO.DirectoryInfo $Path))
            while ($stack.Count -gt 0) {
                $dir = $stack.Pop()
                try { $entries = $dir.GetFileSystemInfos() } catch { continue }
                foreach ($e in $entries) {
                    if ($e -is [System.IO.DirectoryInfo]) {
                        if (($e.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) { $stack.Push($e); $dirs.Add($e) }
                    } else { $files.Add($e) }
                }
            }
        }
        return @{ Files = $files; Dirs = $dirs }
    }

    function Get-FilesSafe([string]$Path) {
        return ,((Get-TreeSafe $Path).Files)
    }

    function Get-FolderSize([string]$Path) {
        $sum = 0
        foreach ($f in (Get-FilesSafe $Path)) { $sum += $f.Length }
        return $sum
    }

    # Deletes files older than $Hours below $Path (skips files in use). Returns freed bytes.
    function Remove-OldFiles([string]$Path, [int]$Hours = 24) {
        if (-not [System.IO.Directory]::Exists($Path)) { return 0 }
        $limit = (Get-Date).AddHours(-$Hours)
        $freed = 0; $n = 0; $skip = 0
        $tree = Get-TreeSafe $Path
        foreach ($f in $tree.Files) {
            if ($Sync.Cancel) { break }
            if ($f.LastWriteTime -gt $limit) { continue }
            try {
                $len = $f.Length
                if ($f.IsReadOnly) { $f.IsReadOnly = $false }
                $f.Delete()
                $freed += $len; $n++
            } catch { $skip++ }
        }
        # remove empty sub-folders (deepest first; junctions are never entered)
        foreach ($d in @($tree.Dirs | Sort-Object { $_.FullName.Length } -Descending)) {
            try { if ($d.LastWriteTime -lt $limit -and $d.GetFileSystemInfos().Length -eq 0) { $d.Delete() } } catch { }
        }
        Log ('{0}: usunięto {1} plików ({2}), pominięto {3} (w użyciu lub nowsze niż {4} h)' -f $Path, $n, (Format-Size $freed), $skip, $Hours) 'out'
        return $freed
    }

    function Get-FreeBytes([string]$Drive = $env:SystemDrive) {
        try { return (New-Object System.IO.DriveInfo $Drive).AvailableFreeSpace } catch { return 0 }
    }

    function Restart-Explorer {
        Get-SessionExplorer | Stop-Process -Force -ErrorAction SilentlyContinue
        for ($i = 0; $i -lt 12; $i++) {
            Start-Sleep -Milliseconds 500
            if ((Get-SessionExplorer).Count -gt 0) { break }
        }
        if ((Get-SessionExplorer).Count -eq 0) { Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') }
        Ok 'Eksplorator Windows został uruchomiony ponownie.'
    }

    function Get-ReportDir {
        if (-not (Test-Path -LiteralPath $ReportRoot)) { New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null }
        return $ReportRoot
    }
}.ToString()

    # =====================================================================
    #  CATEGORIES & ACTIONS (colourful tiles on the first tab)
    #  Code  = runs in a background runspace (helpers from worker prelude)
    #  Ui    = runs on the UI thread instead of Code
    #  Risky = restore point is created first (when the header checkbox is on)
    # =====================================================================
    $Categories = [ordered]@{
        sys  = @{ Name = 'System Windows';               Color = '#3B82F6'; Desc = 'Pliki systemowe, Windows Update, uruchamianie, awarie' }
        net  = @{ Name = 'Sieć i Internet';              Color = '#10B981'; Desc = 'Brak internetu, DNS, Wi-Fi, proxy, karty sieciowe' }
        disk = @{ Name = 'Dysk i wolne miejsce';         Color = '#F97316'; Desc = 'Sprzątanie, TRIM / defragmentacja, stan dysków' }
        mem  = @{ Name = 'Pamięć RAM';                   Color = '#F59E0B'; Desc = 'Zużycie pamięci, plik stronicowania, diagnostyka' }
        perf = @{ Name = 'Wydajność i rozruch';          Color = '#8B5CF6'; Desc = 'Plany zasilania, efekty wizualne, czas uruchamiania' }
        ui   = @{ Name = 'Eksplorator, Start i wygląd';  Color = '#EC4899'; Desc = 'Menu Start, pasek zadań, ikony, menu kontekstowe' }
        dev  = @{ Name = 'Urządzenia, dźwięk i druk';    Color = '#06B6D4'; Desc = 'Drukarki, audio, Bluetooth, sterowniki, bateria' }
        sec  = @{ Name = 'Bezpieczeństwo';               Color = '#EF4444'; Desc = 'Microsoft Defender, zapora, stan ochrony' }
        reg  = @{ Name = 'Rejestr';                      Color = '#84CC16'; Desc = 'Kopie, czyszczenie, historia, defragmentacja' }
    }

    $Actions = @(
        # ------------------------------------------------------------ SYSTEM
        @{ Cat = 'sys'; Title = 'Skanuj pliki systemowe (SFC)'; Tags = @('10–20 min')
           Desc = 'Sprawdza i naprawia uszkodzone lub podmienione pliki systemu Windows.'
           Code = {
               Head 'Kontrola plików systemowych: sfc /scannow'
               $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\sfc.exe') '/scannow' -Unicode
               Log ('Kod wyjścia SFC: {0}. Szczegóły: {1}\Logs\CBS\CBS.log' -f $c, $env:SystemRoot)
               if ($c -eq 0) { Ok 'SFC zakończone.' } else { Warn 'SFC zgłosił problem – jeśli nie zostały naprawione, uruchom "Naprawa obrazu (DISM)" i ponownie SFC.' }
           } }
        @{ Cat = 'sys'; Title = 'Napraw obraz systemu (DISM)'; Tags = @('internet', '10–30 min')
           Desc = 'Naprawia magazyn składników Windows (źródło dla SFC) przy użyciu Windows Update.'
           Code = {
               Head 'DISM /Online /Cleanup-Image /RestoreHealth'
               $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\dism.exe') '/Online /Cleanup-Image /RestoreHealth'
               if ($c -eq 0) { Ok 'Obraz systemu jest sprawny lub został naprawiony.' } else { Err ('DISM zakończył się kodem {0}. Sprawdź połączenie z internetem i log: {1}\Logs\DISM\dism.log' -f $c, $env:SystemRoot) }
           } }
        @{ Cat = 'sys'; Title = 'Pełna naprawa: DISM + SFC'; Tags = @('zalecane', '20–45 min')
           Desc = 'Najskuteczniejsza kolejność: najpierw naprawa obrazu, potem plików systemowych.'
           Code = {
               Head 'Krok 1/2: DISM RestoreHealth'
               $c1 = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\dism.exe') '/Online /Cleanup-Image /RestoreHealth'
               if ($c1 -eq 0) { Ok 'DISM zakończony poprawnie.' } else { Warn ('DISM zwrócił kod {0} – kontynuuję z SFC.' -f $c1) }
               if (Test-Cancel) { return }
               Head 'Krok 2/2: SFC /scannow'
               $c2 = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\sfc.exe') '/scannow' -Unicode
               Log ('Kod wyjścia SFC: {0}' -f $c2)
               Ok 'Pełna naprawa zakończona. Uruchom komputer ponownie, jeśli naprawiono pliki.'
           } }
        @{ Cat = 'sys'; Title = 'Napraw Windows Update'; Tags = @('zalecane'); Risky = $true
           Desc = 'Resetuje usługi i pamięć podręczną aktualizacji (SoftwareDistribution, catroot2).'
           Confirm = 'Usługi Windows Update zostaną zatrzymane, a ich pamięć podręczna wyczyszczona (historia pobranych aktualizacji zostanie odbudowana). Kontynuować?'
           Code = {
               Head 'Reset składników Windows Update'
               $svcs = @('wuauserv', 'UsoSvc', 'bits', 'cryptsvc', 'msiserver')
               foreach ($s in $svcs) { Stop-ServiceSafe $s }
               foreach ($d in @('SoftwareDistribution', 'System32\catroot2')) {
                   $src = Join-Path $env:SystemRoot $d
                   $bak = $src + '.bak'
                   if (-not (Test-Path -LiteralPath $src)) { continue }
                   if (Test-Path -LiteralPath $bak) { Remove-Item -LiteralPath $bak -Recurse -Force -ErrorAction SilentlyContinue }
                   $done = $false
                   for ($i = 0; $i -lt 5 -and -not $done; $i++) {
                       try { Rename-Item -LiteralPath $src -NewName ([System.IO.Path]::GetFileName($bak)) -ErrorAction Stop; $done = $true }
                       catch { foreach ($s in $svcs) { Stop-ServiceSafe $s }; Start-Sleep -Seconds 2 }
                   }
                   if ($done) { Ok ('Zmieniono nazwę: {0} -> {1}' -f $src, $bak) } else { Warn ('Nie udało się zmienić nazwy {0} (plik w użyciu). Spróbuj po restarcie.' -f $src) }
               }
               foreach ($s in @('cryptsvc', 'bits', 'wuauserv', 'UsoSvc')) { Start-ServiceSafe $s }
               Ok 'Windows Update zresetowany. Sprawdź aktualizacje w Ustawieniach.'
           } }
        @{ Cat = 'sys'; Title = 'Sprawdź dysk systemowy (CHKDSK)'; Tags = @('bez restartu')
           Desc = 'Skanuje system plików dysku C: w trybie online i zgłasza błędy.'
           Code = {
               Head ('chkdsk {0} /scan' -f $env:SystemDrive)
               $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\chkdsk.exe') ($env:SystemDrive + ' /scan')
               switch ($c) {
                   0 { Ok 'Nie znaleziono błędów systemu plików.' }
                   1 { Ok 'Znalezione błędy zostały naprawione.' }
                   default { Warn 'Wykryto problemy wymagające naprawy offline – użyj "Zaplanuj naprawę dysku przy restarcie".' }
               }
           } }
        @{ Cat = 'sys'; Title = 'Zaplanuj naprawę dysku przy restarcie'; Tags = @('restart')
           Desc = 'Przy następnym uruchomieniu Windows sprawdzi i naprawi dysk C: (chkdsk /f).'
           Confirm = 'Przy następnym uruchomieniu komputera zostanie wykonane sprawdzenie dysku systemowego (może potrwać kilka–kilkanaście minut). Kontynuować?'
           Code = {
               $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\fsutil.exe') ('dirty set ' + $env:SystemDrive)
               if ($c -eq 0) { Ok 'Naprawa dysku zostanie wykonana przy najbliższym restarcie.' } else { Err 'Nie udało się zaplanować sprawdzenia dysku.' }
           } }
        @{ Cat = 'sys'; Title = 'Przebuduj indeks wyszukiwania'; Tags = @()
           Desc = 'Gdy wyszukiwarka w menu Start lub Eksploratorze nie znajduje plików.'
           Code = {
               Stop-ServiceSafe 'WSearch'
               Set-Reg 'HKLM:\SOFTWARE\Microsoft\Windows Search' 'SetupCompletedSuccessfully' 0 'DWord'
               Start-ServiceSafe 'WSearch'
               Ok 'Indeks zostanie odbudowany w tle (może to potrwać do kilku godzin, wyniki pojawiają się stopniowo).'
           } }
        @{ Cat = 'sys'; Title = 'Napraw Microsoft Store'; Tags = @()
           Desc = 'Czyści pamięć podręczną sklepu (wsreset) – pomaga, gdy aplikacje nie pobierają się.'
           Code = {
               Log 'Uruchamiam wsreset.exe – po zakończeniu otworzy się Microsoft Store.'
               try { Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\wsreset.exe') -Wait -ErrorAction Stop; Ok 'Pamięć podręczna sklepu wyczyszczona.' }
               catch { Err $_.Exception.Message }
           } }
        @{ Cat = 'sys'; Title = 'Synchronizuj zegar systemowy'; Tags = @('internet')
           Desc = 'Zła godzina powoduje błędy certyfikatów, logowania i stron HTTPS.'
           Code = {
               Start-ServiceSafe 'w32time'
               $w = Join-Path $env:SystemRoot 'System32\w32tm.exe'
               $c = Invoke-Cmd $w '/resync /force'
               if ($c -ne 0) {
                   Warn 'Pierwsza próba nieudana – ustawiam serwery czasu i ponawiam.'
                   [void](Invoke-Cmd $w '/config /manualpeerlist:"time.windows.com,0x9 pool.ntp.org,0x9" /syncfromflags:manual /update')
                   Restart-Service -Name w32time -Force -ErrorAction SilentlyContinue
                   Start-Sleep -Seconds 2
                   $c = Invoke-Cmd $w '/resync /force'
               }
               if ($c -eq 0) { Ok ('Zegar zsynchronizowany: {0}' -f (Get-Date)) } else { Err 'Synchronizacja nieudana – sprawdź połączenie z internetem.' }
           } }
        @{ Cat = 'sys'; Title = 'Szybkie uruchamianie: wł. / wył.'; Tags = @('przełącznik')
           Desc = 'Wyłączenie często naprawia problemy z zamykaniem, sterownikami i dual-bootem.'
           Confirm = {
               $v = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -ErrorAction SilentlyContinue).HiberbootEnabled
               if ($v -eq 0) { 'Szybkie uruchamianie jest teraz WYŁĄCZONE. Włączyć je?' }
               else { "Szybkie uruchamianie jest teraz WŁĄCZONE.`n`nWyłączenie często rozwiązuje problemy z zamykaniem systemu, sterownikami, Wi-Fi po wybudzeniu i dual-bootem (start trwa wtedy nieco dłużej).`n`nWyłączyć?" }
           }
           Code = {
               $k = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power'
               $v = Get-RegValue $k 'HiberbootEnabled'
               if ($v -eq 0) { $new = 1 } else { $new = 0 }
               Set-Reg $k 'HiberbootEnabled' $new 'DWord'
               if ($new -eq 1) {
                   Ok 'Szybkie uruchamianie WŁĄCZONE.'
                   if ((Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' 'HibernateEnabled') -eq 0) { Warn 'Hibernacja jest wyłączona – szybkie uruchamianie zadziała dopiero po jej włączeniu.' }
               } else { Ok 'Szybkie uruchamianie WYŁĄCZONE.' }
           } }
        @{ Cat = 'sys'; Title = 'Raport awarii (BSOD i zawieszenia)'; Tags = @('diagnostyka')
           Desc = 'Niebieskie ekrany, nagłe wyłączenia i awarie aplikacji z ostatnich 30 dni.'
           Code = {
               $since = (Get-Date).AddDays(-30)
               Head 'Niebieskie ekrany (BugCheck) – 30 dni'
               $bsod = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 1001; StartTime = $since } -ErrorAction SilentlyContinue |
                   Where-Object { $_.ProviderName -match 'WER-SystemErrorReporting|BugCheck' })
               if ($bsod.Count -eq 0) { Ok 'Brak zarejestrowanych niebieskich ekranów.' }
               foreach ($e in $bsod) { Warn ('{0:yyyy-MM-dd HH:mm}  {1}' -f $e.TimeCreated, (($e.Message -split "`n")[0]).Trim()) }
               $kp = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = 41; StartTime = $since } -ErrorAction SilentlyContinue)
               if ($kp.Count -gt 0) { Warn ('Nieoczekiwane wyłączenia (Kernel-Power 41): {0}, ostatnie: {1:yyyy-MM-dd HH:mm}' -f $kp.Count, $kp[0].TimeCreated) } else { Ok 'Brak nieoczekiwanych wyłączeń zasilania.' }
               $dumps = @(Get-ChildItem -Path (Join-Path $env:SystemRoot 'Minidump') -Filter *.dmp -ErrorAction SilentlyContinue)
               if ($dumps.Count) { Log ('Zrzuty pamięci w {0}\Minidump: {1}' -f $env:SystemRoot, $dumps.Count) }
               Head 'Najczęściej zawieszające się aplikacje – 30 dni'
               $app = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Id = 1000, 1002; StartTime = $since } -ErrorAction SilentlyContinue)
               if ($app.Count -eq 0) { Ok 'Brak awarii aplikacji.' }
               else {
                   $app | ForEach-Object { if ($_.Properties.Count) { [string]$_.Properties[0].Value } } |
                       Group-Object | Sort-Object Count -Descending | Select-Object -First 10 |
                       ForEach-Object { Log ('{0,4} x  {1}' -f $_.Count, $_.Name) 'out' }
               }
           } }
        @{ Cat = 'sys'; Title = 'Utwórz punkt przywracania'; Tags = @('bezpieczeństwo')
           Desc = 'Zapisuje stan systemu, do którego można wrócić w razie problemów.'
           Code = { [void](New-RestorePointSafe 'RegDoktor - punkt reczny') } }

        # ------------------------------------------------------------ NETWORK
        @{ Cat = 'net'; Title = 'Diagnostyka połączenia'; Tags = @('zacznij tutaj')
           Desc = 'Sprawdza kartę, router, internet, DNS i HTTPS – i podpowiada, co naprawić.'
           Code = {
               Head 'Konfiguracja sieci'
               $cfgs = @(Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.NetAdapter.Status -eq 'Up' })
               if ($cfgs.Count -eq 0) { Err 'Brak aktywnej karty sieciowej (kabel odłączony / Wi-Fi rozłączone / karta wyłączona).'; return }
               $gws = @()
               foreach ($c in $cfgs) {
                   Log ('{0}  ({1})' -f $c.InterfaceAlias, $c.InterfaceDescription)
                   Log ('   IPv4: {0}' -f ((@($c.IPv4Address) | ForEach-Object { $_.IPAddress }) -join ', ')) 'out'
                   $gw = @($c.IPv4DefaultGateway | ForEach-Object { $_.NextHop })
                   $gws += $gw
                   Log ('   Brama: {0}' -f ($gw -join ', ')) 'out'
                   Log ('   DNS: {0}' -f ((@($c.DNSServer) | Where-Object { $_.AddressFamily -eq 2 } | ForEach-Object { $_.ServerAddresses }) -join ', ')) 'out'
                   if (@($c.IPv4Address | Where-Object { $_.IPAddress -like '169.254.*' }).Count) { Warn '   Adres 169.254.x.x – karta nie dostała adresu z routera (DHCP).' }
               }
               $ping = New-Object System.Net.NetworkInformation.Ping
               function Test-Ping([string]$h) { try { $r = $ping.Send($h, 2000); if ($r.Status -eq 'Success') { return $r.RoundtripTime } } catch { }; return -1 }
               Head 'Testy'
               $gwOk = $false
               foreach ($g in ($gws | Where-Object { $_ } | Select-Object -Unique)) {
                   $t = Test-Ping $g
                   if ($t -ge 0) { Ok ('Router {0}: {1} ms' -f $g, $t); $gwOk = $true } else { Warn ('Router {0} nie odpowiada na ping (część routerów to blokuje).' -f $g) }
               }
               $netOk = $false
               foreach ($h in @('1.1.1.1', '8.8.8.8')) { $t = Test-Ping $h; if ($t -ge 0) { Ok ('Internet ({0}): {1} ms' -f $h, $t); $netOk = $true } else { Warn ('Brak odpowiedzi z {0}' -f $h) } }
               $dnsOk = $false
               try { $ips = [System.Net.Dns]::GetHostAddresses('www.microsoft.com'); Ok ('DNS: www.microsoft.com -> {0}' -f ($ips[0].IPAddressToString)); $dnsOk = $true } catch { Err ('DNS nie działa: {0}' -f $_.Exception.Message) }
               $httpOk = $false
               try {
                   $req = [System.Net.WebRequest]::Create('http://www.msftconnecttest.com/connecttest.txt')
                   $req.Timeout = 6000
                   $resp = $req.GetResponse(); $sr = New-Object System.IO.StreamReader($resp.GetResponseStream()); $body = $sr.ReadToEnd(); $resp.Close()
                   if ($body -match 'Microsoft Connect Test') { Ok 'HTTP: test łączności Microsoft OK'; $httpOk = $true } else { Warn 'HTTP: odpowiedź nietypowa – możliwa strona logowania (hotel/lotnisko) lub proxy.' }
               } catch { Warn ('HTTP: {0}' -f $_.Exception.Message) }
               try {
                   $tcp = New-Object System.Net.Sockets.TcpClient
                   $ar = $tcp.BeginConnect('www.microsoft.com', 443, $null, $null)
                   if ($ar.AsyncWaitHandle.WaitOne(4000) -and $tcp.Connected) { Ok 'HTTPS (port 443): połączenie OK' } else { Warn 'HTTPS (port 443): brak połączenia' }
                   $tcp.Close()
               } catch { Warn ('HTTPS: {0}' -f $_.Exception.Message) }
               Head 'Wniosek'
               if ($netOk -and $dnsOk -and $httpOk) { Ok 'Połączenie z internetem działa poprawnie.' }
               elseif ($netOk -and -not $dnsOk) { Warn 'Internet jest, ale nie działa DNS -> użyj "Wyczyść pamięć DNS" lub "DNS: Cloudflare + Google".' }
               elseif ($netOk -and $dnsOk -and -not $httpOk) { Warn 'DNS i internet działają, ale HTTP nie -> sprawdź "Resetuj ustawienia proxy" i zaporę/antywirusa.' }
               elseif (-not $netOk -and $gwOk) { Warn 'Router odpowiada, ale internet nie -> problem u dostawcy lub w routerze (uruchom go ponownie).' }
               else { Warn 'Brak łączności z routerem -> "Zrestartuj karty sieciowe", a potem "Pełna naprawa sieci".' }
           } }
        @{ Cat = 'net'; Title = 'Wyczyść pamięć DNS'; Tags = @('szybkie')
           Desc = 'Gdy strony się nie otwierają lub wskazują stary adres.'
           Code = {
               [void](Invoke-Cmd (Join-Path $env:SystemRoot 'System32\ipconfig.exe') '/flushdns')
               Clear-DnsClientCache -ErrorAction SilentlyContinue
               Ok 'Pamięć podręczna DNS wyczyszczona.'
           } }
        @{ Cat = 'net'; Title = 'Odnów adres IP'; Tags = @('chwilowe rozłączenie')
           Desc = 'Pobiera nowy adres z routera (DHCP) – np. przy adresie 169.254.x.x.'
           Confirm = 'Połączenie sieciowe zostanie na chwilę przerwane. Kontynuować?'
           Code = {
               $ip = Join-Path $env:SystemRoot 'System32\ipconfig.exe'
               [void](Invoke-Cmd $ip '/release')
               $c = Invoke-Cmd $ip '/renew'
               if ($c -eq 0) { Ok 'Adres IP odnowiony.' } else { Warn 'Odnowienie nie powiodło się na wszystkich kartach (to normalne dla kart rozłączonych).' }
           } }
        @{ Cat = 'net'; Title = 'Resetuj Winsock'; Tags = @('restart')
           Desc = 'Naprawia uszkodzony katalog Winsock (np. po usunięciu VPN lub antywirusa).'
           Code = { $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\netsh.exe') 'winsock reset'; if ($c -eq 0) { Ok 'Winsock zresetowany – wymagany restart.' } } }
        @{ Cat = 'net'; Title = 'Resetuj stos TCP/IP'; Tags = @('restart')
           Desc = 'Przywraca domyślne ustawienia protokołów IPv4 i IPv6.'
           Code = {
               $n = Join-Path $env:SystemRoot 'System32\netsh.exe'
               [void](Invoke-Cmd $n 'int ip reset')
               [void](Invoke-Cmd $n 'int ipv6 reset')
               Ok 'Stos TCP/IP zresetowany – wymagany restart.'
           } }
        @{ Cat = 'net'; Title = 'Pełna naprawa sieci'; Tags = @('zalecane', 'restart'); Risky = $true
           Desc = 'DNS + ARP + NetBIOS + Winsock + TCP/IP + proxy + nowy adres IP – w jednym kroku.'
           Confirm = 'Zostaną zresetowane ustawienia sieciowe (Winsock, TCP/IP, proxy WinHTTP), a połączenie na chwilę zerwane. Statyczne adresy IP mogą wymagać ponownego wpisania. Kontynuować?'
           Code = {
               $n = Join-Path $env:SystemRoot 'System32\netsh.exe'
               $ip = Join-Path $env:SystemRoot 'System32\ipconfig.exe'
               $nb = Join-Path $env:SystemRoot 'System32\nbtstat.exe'
               [void](Invoke-Cmd $ip '/flushdns')
               [void](Invoke-Cmd $n 'interface ip delete arpcache')
               [void](Invoke-Cmd $nb '-R')
               [void](Invoke-Cmd $nb '-RR')
               [void](Invoke-Cmd $n 'winsock reset')
               [void](Invoke-Cmd $n 'int ip reset')
               [void](Invoke-Cmd $n 'int ipv6 reset')
               [void](Invoke-Cmd $n 'winhttp reset proxy')
               [void](Invoke-Cmd $ip '/release')
               [void](Invoke-Cmd $ip '/renew')
               Ok 'Naprawa sieci zakończona – uruchom komputer ponownie.'
           } }
        @{ Cat = 'net'; Title = 'Zrestartuj karty sieciowe'; Tags = @('chwilowe rozłączenie')
           Desc = 'Wyłącza i włącza fizyczne karty Ethernet / Wi-Fi.'
           Confirm = 'Wszystkie fizyczne karty sieciowe zostaną na chwilę wyłączone. Kontynuować?'
           Code = {
               $ad = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'Disabled' })
               if ($ad.Count -eq 0) { Warn 'Nie znaleziono aktywnych kart.'; return }
               foreach ($a in $ad) {
                   Log ('Restart: {0} ({1})' -f $a.Name, $a.InterfaceDescription)
                   try { Restart-NetAdapter -Name $a.Name -Confirm:$false -ErrorAction Stop; Ok ('{0} uruchomiona ponownie' -f $a.Name) } catch { Err $_.Exception.Message }
               }
           } }
        @{ Cat = 'net'; Title = 'Wyczyść ARP i NetBIOS'; Tags = @('szybkie')
           Desc = 'Gdy komputery w sieci lokalnej się nie widzą lub drukarka sieciowa znika.'
           Code = {
               [void](Invoke-Cmd (Join-Path $env:SystemRoot 'System32\netsh.exe') 'interface ip delete arpcache')
               [void](Invoke-Cmd (Join-Path $env:SystemRoot 'System32\nbtstat.exe') '-R')
               [void](Invoke-Cmd (Join-Path $env:SystemRoot 'System32\nbtstat.exe') '-RR')
               Ok 'Pamięć ARP i NetBIOS wyczyszczona.'
           } }
        @{ Cat = 'net'; Title = 'Resetuj ustawienia proxy'; Tags = @()
           Desc = 'Gdy przeglądarki nie łączą się, a sieć działa (np. po złośliwym oprogramowaniu).'
           Code = {
               Backup-RegKey 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings' 'InternetSettings'
               [void](Invoke-Cmd (Join-Path $env:SystemRoot 'System32\netsh.exe') 'winhttp reset proxy')
               $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
               Set-Reg $k 'ProxyEnable' 0 'DWord'
               Remove-ItemProperty -LiteralPath $k -Name 'ProxyServer' -ErrorAction SilentlyContinue
               Remove-ItemProperty -LiteralPath $k -Name 'AutoConfigURL' -ErrorAction SilentlyContinue
               Ok 'Proxy wyłączone (WinHTTP i ustawienia użytkownika).'
           } }
        @{ Cat = 'net'; Title = 'DNS: Cloudflare + Google'; Tags = @()
           Desc = 'Ustawia szybkie, publiczne serwery DNS 1.1.1.1 i 8.8.8.8 na aktywnych kartach.'
           Code = {
               $ad = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' })
               foreach ($a in $ad) {
                   try { Set-DnsClientServerAddress -InterfaceIndex $a.ifIndex -ServerAddresses @('1.1.1.1', '8.8.8.8') -ErrorAction Stop; Ok ('{0}: DNS = 1.1.1.1, 8.8.8.8' -f $a.Name) } catch { Err $_.Exception.Message }
               }
               Clear-DnsClientCache -ErrorAction SilentlyContinue
           } }
        @{ Cat = 'net'; Title = 'DNS: automatyczne (z routera)'; Tags = @()
           Desc = 'Przywraca pobieranie serwerów DNS z routera (DHCP).'
           Code = {
               $ad = @(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' })
               foreach ($a in $ad) {
                   try { Set-DnsClientServerAddress -InterfaceIndex $a.ifIndex -ResetServerAddresses -ErrorAction Stop; Ok ('{0}: DNS automatyczne' -f $a.Name) } catch { Err $_.Exception.Message }
               }
               Clear-DnsClientCache -ErrorAction SilentlyContinue
           } }
        @{ Cat = 'net'; Title = 'Raport Wi-Fi'; Tags = @('raport')
           Desc = 'Szczegółowy raport HTML: rozłączenia, błędy i sygnał z ostatnich 3 dni.'
           Code = {
               $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\netsh.exe') 'wlan show wlanreport'
               $f = Join-Path $env:ProgramData 'Microsoft\Windows\WlanReport\wlan-report-latest.html'
               if (Test-Path -LiteralPath $f) { $dst = Join-Path (Get-ReportDir) 'raport-wifi.html'; Copy-Item -LiteralPath $f -Destination $dst -Force; Open-Unelevated $dst; Ok ('Raport: {0}' -f $dst) }
               else { Warn 'Nie wygenerowano raportu (brak karty Wi-Fi?).' }
           } }
        @{ Cat = 'net'; Title = 'Pełny reset sieci Windows'; Tags = @('restart', 'ostateczność'); Risky = $true
           Desc = 'Usuwa i instaluje ponownie wszystkie karty sieciowe (jak w Ustawieniach).'
           Confirm = "Wszystkie karty sieciowe (także VPN i wirtualne) zostaną usunięte i zainstalowane ponownie, zapisane ustawienia kart zostaną utracone. Hasła Wi-Fi pozostają.`n`nTo ostateczność – użyj, gdy nic innego nie pomogło. Kontynuować?"
           Code = { $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\netcfg.exe') '-d'; if ($c -eq 0) { Ok 'Reset sieci wykonany – uruchom komputer ponownie.' } else { Err ('netcfg zwrócił kod {0}' -f $c) } } }

        # ------------------------------------------------------------ DISK
        @{ Cat = 'disk'; Title = 'Co zajmuje miejsce?'; Tags = @('analiza')
           Desc = 'Rozmiar typowych "pożeraczy" miejsca: Temp, aktualizacje, Kosz, Windows.old…'
           Code = {
               $d = New-Object System.IO.DriveInfo $env:SystemDrive
               Head ('Dysk {0}: {1} wolne z {2}' -f $env:SystemDrive, (Format-Size $d.AvailableFreeSpace), (Format-Size $d.TotalSize))
               $items = @(
                   @('Pliki tymczasowe użytkownika', $env:TEMP),
                   @('Pliki tymczasowe systemu', (Join-Path $env:SystemRoot 'Temp')),
                   @('Pobrane aktualizacje Windows', (Join-Path $env:SystemRoot 'SoftwareDistribution\Download')),
                   @('Poprzednia instalacja (Windows.old)', (Join-Path $env:SystemDrive 'Windows.old')),
                   @('Kosz', (Join-Path $env:SystemDrive '$Recycle.Bin')),
                   @('Folder Pobrane', (Join-Path $env:USERPROFILE 'Downloads')),
                   @('Zrzuty awarii aplikacji', (Join-Path $env:LOCALAPPDATA 'CrashDumps')),
                   @('Raporty błędów (WER)', (Join-Path $env:ProgramData 'Microsoft\Windows\WER')),
                   @('Optymalizacja dostarczania', (Join-Path $env:SystemRoot 'ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization')),
                   @('Instalatory MSI (NIE usuwać ręcznie)', (Join-Path $env:SystemRoot 'Installer'))
               )
               $rows = @()
               $i = 0
               foreach ($it in $items) {
                   $i++; Progress ($i * 100 / ($items.Count + 1)) ('Liczenie: ' + $it[0])
                   if (Test-Cancel) { return }
                   $rows += New-Object PSObject -Property @{ Name = $it[0]; Size = [double](Get-FolderSize $it[1]); Path = $it[1] }
               }
               foreach ($f in @('hiberfil.sys', 'pagefile.sys', 'swapfile.sys')) {
                   $p = Join-Path ($env:SystemDrive + '\') $f
                   try { $fi = Get-Item -LiteralPath $p -Force -ErrorAction Stop; $rows += New-Object PSObject -Property @{ Name = $f; Size = [double]$fi.Length; Path = $p } } catch { }
               }
               $md = Join-Path $env:SystemRoot 'MEMORY.DMP'
               if (Test-Path -LiteralPath $md) { $rows += New-Object PSObject -Property @{ Name = 'Zrzut pamięci MEMORY.DMP'; Size = [double](Get-Item -LiteralPath $md -Force).Length; Path = $md } }
               foreach ($r in ($rows | Sort-Object Size -Descending)) {
                   if ($r.Size -le 0) { continue }
                   $kind = 'out'; if ($r.Size -gt 5GB) { $kind = 'warn' }
                   Log ('{0,10}   {1}   ({2})' -f (Format-Size $r.Size), $r.Name, $r.Path) $kind
               }
               Log 'Wskazówka: "Oczyszczanie dysku" usuwa bezpiecznie większość z tych pozycji (poza folderem Pobrane).'
           } }
        @{ Cat = 'disk'; Title = 'Usuń pliki tymczasowe'; Tags = @('bezpieczne')
           Desc = 'Temp użytkownika i systemu, zrzuty awarii, raporty błędów (starsze niż 24 h).'
           Code = {
               $free0 = Get-FreeBytes
               $sum = 0
               foreach ($p in @($env:TEMP, (Join-Path $env:SystemRoot 'Temp'), (Join-Path $env:LOCALAPPDATA 'CrashDumps'),
                       (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportArchive'), (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportQueue'),
                       (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\INetCache\IE'), (Join-Path $env:SystemRoot 'Logs\CBS\Temp'))) {
                   if (Test-Cancel) { break }
                   $sum += Remove-OldFiles $p 24
               }
               Ok ('Zwolniono około {0}. (Wolne miejsce: {1} -> {2})' -f (Format-Size $sum), (Format-Size $free0), (Format-Size (Get-FreeBytes)))
           } }
        @{ Cat = 'disk'; Title = 'Oczyszczanie dysku (systemowe)'; Tags = @('5–30 min')
           Desc = 'Wbudowane narzędzie Windows: stare aktualizacje, logi, miniatury, Kosz… (bez Pobranych).'
           Confirm = 'Uruchomić Oczyszczanie dysku dla wszystkich kategorii (poza folderem Pobrane, plikami instalacyjnymi ESD i poprzednią instalacją Windows)?'
           Code = {
               $base = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
               $skip = @('DownloadsFolder', 'Previous Installations', 'Windows ESD installation files')
               foreach ($k in @(Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue)) {
                   if ($skip -contains $k.PSChildName) { Remove-ItemProperty -LiteralPath $k.PSPath -Name 'StateFlags0777' -ErrorAction SilentlyContinue; continue }
                   New-ItemProperty -LiteralPath $k.PSPath -Name 'StateFlags0777' -Value 2 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
               }
               $free0 = Get-FreeBytes
               Progress -1 'Oczyszczanie dysku pracuje (osobne okno)…'
               Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\cleanmgr.exe') -ArgumentList '/sagerun:777' -Wait
               $gain = (Get-FreeBytes) - $free0
               Ok ('Oczyszczanie zakończone. Zwolniono około {0}.' -f (Format-Size ([Math]::Max(0, $gain))))
           } }
        @{ Cat = 'disk'; Title = 'Wyczyść składniki (WinSxS)'; Tags = @('5–20 min')
           Desc = 'Usuwa zastąpione wersje składników po aktualizacjach (DISM StartComponentCleanup).'
           Code = {
               $free0 = Get-FreeBytes
               $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\dism.exe') '/Online /Cleanup-Image /StartComponentCleanup'
               if ($c -eq 0) { Ok ('Zakończono. Zwolniono około {0}.' -f (Format-Size ([Math]::Max(0, (Get-FreeBytes) - $free0)))) } else { Err ('DISM zwrócił kod {0}' -f $c) }
           } }
        @{ Cat = 'disk'; Title = 'Usuń Windows.old'; Tags = @('nieodwracalne')
           Desc = 'Kasuje poprzednią instalację Windows (często 10–30 GB).'
           Confirm = 'Po usunięciu Windows.old NIE będzie można wrócić do poprzedniej wersji Windows. Kontynuować?'
           Code = {
               if (-not (Test-Path -LiteralPath (Join-Path $env:SystemDrive 'Windows.old'))) { Ok 'Brak folderu Windows.old – nie ma czego usuwać.'; return }
               $base = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
               foreach ($k in @(Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue)) {
                   if ($k.PSChildName -eq 'Previous Installations') { New-ItemProperty -LiteralPath $k.PSPath -Name 'StateFlags0778' -Value 2 -PropertyType DWord -Force | Out-Null }
                   else { Remove-ItemProperty -LiteralPath $k.PSPath -Name 'StateFlags0778' -ErrorAction SilentlyContinue }
               }
               $free0 = Get-FreeBytes
               Progress -1 'Usuwanie Windows.old (osobne okno)…'
               Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\cleanmgr.exe') -ArgumentList '/sagerun:778' -Wait
               Ok ('Zwolniono około {0}.' -f (Format-Size ([Math]::Max(0, (Get-FreeBytes) - $free0))))
           } }
        @{ Cat = 'disk'; Title = 'Opróżnij Kosz'; Tags = @('nieodwracalne')
           Desc = 'Trwale usuwa zawartość Kosza na wszystkich dyskach.'
           Confirm = 'Trwale usunąć zawartość Kosza?'
           Code = {
               $free0 = Get-FreeBytes
               try { Clear-RecycleBin -Force -ErrorAction Stop } catch { }
               Ok ('Kosz opróżniony. Zwolniono na {0} około {1}.' -f $env:SystemDrive, (Format-Size ([Math]::Max(0, (Get-FreeBytes) - $free0))))
           } }
        @{ Cat = 'disk'; Title = 'Wyczyść Optymalizację dostarczania'; Tags = @('bezpieczne')
           Desc = 'Pamięć podręczna aktualizacji współdzielonych z innymi komputerami.'
           Code = {
               $free0 = Get-FreeBytes
               try { Delete-DeliveryOptimizationCache -Force -ErrorAction Stop; Ok ('Wyczyszczono. Zwolniono około {0}.' -f (Format-Size ([Math]::Max(0, (Get-FreeBytes) - $free0)))) }
               catch { Err $_.Exception.Message }
           } }
        @{ Cat = 'disk'; Title = 'Optymalizuj dyski (TRIM / defrag)'; Tags = @('zalecane')
           Desc = 'SSD: polecenie TRIM, HDD: defragmentacja – Windows dobiera metodę sam.'
           Code = {
               $vols = @(Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter -and $_.DriveType -eq 'Fixed' -and ($_.FileSystem -eq 'NTFS' -or $_.FileSystem -eq 'ReFS') })
               foreach ($v in $vols) {
                   if (Test-Cancel) { break }
                   $media = 'nieznany'
                   try {
                       $part = Get-Partition -DriveLetter $v.DriveLetter -ErrorAction Stop
                       $pd = Get-PhysicalDisk -ErrorAction SilentlyContinue | Where-Object { $_.DeviceId -eq [string]$part.DiskNumber } | Select-Object -First 1
                       if ($pd) { $media = [string]$pd.MediaType }
                   } catch { }
                   Head ('Dysk {0}: (typ: {1})' -f $v.DriveLetter, $media)
                   Progress -1 ('Optymalizacja {0}:…' -f $v.DriveLetter)
                   try {
                       $out = Optimize-Volume -DriveLetter $v.DriveLetter -Verbose -ErrorAction Stop 4>&1
                       foreach ($o in $out) { $s = [string]$o; if ($s.Trim()) { Log $s 'out' } }
                       Ok ('Dysk {0}: zoptymalizowany' -f $v.DriveLetter)
                   } catch { Warn ('Dysk {0}: {1}' -f $v.DriveLetter, $_.Exception.Message) }
               }
           } }
        @{ Cat = 'disk'; Title = 'Stan dysków (SMART)'; Tags = @('diagnostyka')
           Desc = 'Kondycja, temperatura, zużycie SSD i błędy odczytu/zapisu.'
           Code = {
               foreach ($d in @(Get-PhysicalDisk -ErrorAction SilentlyContinue)) {
                   $kind = 'ok'; if ([string]$d.HealthStatus -ne 'Healthy') { $kind = 'err' }
                   Log ('{0}  [{1}, {2}, {3}]  Stan: {4}' -f $d.FriendlyName, $d.MediaType, $d.BusType, (Format-Size $d.Size), $d.HealthStatus) $kind
                   try {
                       $rc = $d | Get-StorageReliabilityCounter -ErrorAction Stop
                       $parts = @()
                       if ($null -ne $rc.Temperature -and $rc.Temperature -gt 0) { $parts += ('temperatura {0}°C' -f $rc.Temperature) }
                       if ($null -ne $rc.Wear) { $parts += ('zużycie {0}%' -f $rc.Wear) }
                       if ($null -ne $rc.PowerOnHours) { $parts += ('godziny pracy {0}' -f $rc.PowerOnHours) }
                       if ($null -ne $rc.ReadErrorsTotal) { $parts += ('błędy odczytu {0}' -f $rc.ReadErrorsTotal) }
                       if ($null -ne $rc.WriteErrorsTotal) { $parts += ('błędy zapisu {0}' -f $rc.WriteErrorsTotal) }
                       if ($parts.Count) { Log ('   ' + ($parts -join ', ')) 'out' } else { Log '   (dysk nie udostępnia liczników SMART)' 'out' }
                   } catch { Log '   (liczniki SMART niedostępne)' 'out' }
               }
           } }
        @{ Cat = 'disk'; Title = 'Hibernacja: wł. / wył.'; Tags = @('przełącznik')
           Desc = 'Wyłączenie usuwa plik hiberfil.sys (kilka GB), ale też wyłącza szybkie uruchamianie.'
           Confirm = {
               $v = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' -ErrorAction SilentlyContinue).HibernateEnabled
               if ($v -eq 0) { 'Hibernacja jest teraz WYŁĄCZONA. Włączyć ją?' } else { "Hibernacja jest teraz WŁĄCZONA.`n`nWyłączenie zwolni miejsce zajmowane przez hiberfil.sys, ale wyłączy też hibernację i szybkie uruchamianie. Wyłączyć?" }
           }
           Code = {
               $v = Get-RegValue 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' 'HibernateEnabled'
               $pc = Join-Path $env:SystemRoot 'System32\powercfg.exe'
               if ($v -eq 0) { $c = Invoke-Cmd $pc '/hibernate on'; if ($c -eq 0) { Ok 'Hibernacja WŁĄCZONA.' } }
               else {
                   $free0 = Get-FreeBytes
                   $c = Invoke-Cmd $pc '/hibernate off'
                   if ($c -eq 0) { Ok ('Hibernacja WYŁĄCZONA. Zwolniono około {0}.' -f (Format-Size ([Math]::Max(0, (Get-FreeBytes) - $free0)))) }
               }
           } }

        # ------------------------------------------------------------ MEMORY
        @{ Cat = 'mem'; Title = 'Raport pamięci RAM'; Tags = @('diagnostyka')
           Desc = 'Zużycie RAM, pamięć zadeklarowana, plik stronicowania i najbardziej "pamięciożerne" procesy.'
           Code = {
               $os = Get-CimInstance Win32_OperatingSystem
               $tot = [double]$os.TotalVisibleMemorySize * 1KB; $free = [double]$os.FreePhysicalMemory * 1KB
               $vt = [double]$os.TotalVirtualMemorySize * 1KB; $vf = [double]$os.FreeVirtualMemory * 1KB
               Head 'Pamięć'
               $kind = 'ok'; if (($tot - $free) / $tot -gt 0.85) { $kind = 'warn' }
               Log ('RAM: zajęte {0} z {1} ({2:N0}%), dostępne {3}' -f (Format-Size ($tot - $free)), (Format-Size $tot), (100 * ($tot - $free) / $tot), (Format-Size $free)) $kind
               Log ('Pamięć zadeklarowana (commit): {0} z limitu {1}' -f (Format-Size ($vt - $vf)), (Format-Size $vt)) 'out'
               foreach ($pf in @(Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue)) {
                   Log ('Plik stronicowania {0}: rozmiar {1} MB, użyte {2} MB, szczyt {3} MB' -f $pf.Name, $pf.AllocatedBaseSize, $pf.CurrentUsage, $pf.PeakUsage) 'out'
               }
               $cs = Get-CimInstance Win32_ComputerSystem
               Log ('Plik stronicowania zarządzany automatycznie: {0}' -f $(if ($cs.AutomaticManagedPagefile) { 'tak' } else { 'NIE' })) 'out'
               Head 'Procesy zajmujące najwięcej pamięci (suma dla nazwy)'
               Get-Process | Group-Object ProcessName | ForEach-Object {
                   New-Object PSObject -Property @{ Name = $_.Name; Count = $_.Count; Mem = [double](($_.Group | Measure-Object WorkingSet64 -Sum).Sum) }
               } | Sort-Object Mem -Descending | Select-Object -First 12 | ForEach-Object {
                   Log ('{0,10}  {1}{2}' -f (Format-Size $_.Mem), $_.Name, $(if ($_.Count -gt 1) { ' (x' + $_.Count + ')' } else { '' })) 'out'
               }
           } }
        @{ Cat = 'mem'; Title = 'Zwolnij pamięć podręczną (standby)'; Tags = @('szybkie')
           Desc = 'Opróżnia listę standby. Pomaga głównie w grach i przy wyciekach – Windows i tak zarządza nią sam.'
           Code = {
               $before = [double](Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory * 1KB
               $r = [RegDoktor.Native]::MemoryCommand(4)
               Start-Sleep -Milliseconds 500
               $after = [double](Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory * 1KB
               if ($r -eq 0) { Ok ('Lista standby opróżniona. Wolna pamięć: {0} -> {1}' -f (Format-Size $before), (Format-Size $after)) }
               else { Err ('Nie udało się (NTSTATUS 0x{0:X8}).' -f $r) }
           } }
        @{ Cat = 'mem'; Title = 'Top procesów: CPU i RAM'; Tags = @('diagnostyka')
           Desc = 'Co w tej chwili obciąża procesor i pamięć (pomiar 2 s).'
           Code = {
               $t0 = @{}
               foreach ($p in Get-Process) { try { $t0[$p.Id] = $p.TotalProcessorTime.TotalMilliseconds } catch { } }
               Progress -1 'Pomiar obciążenia (2 s)…'
               Start-Sleep -Seconds 2
               $cores = [Environment]::ProcessorCount
               $rows = foreach ($p in Get-Process) {
                   $cpu = 0.0
                   try { if ($t0.ContainsKey($p.Id)) { $cpu = ($p.TotalProcessorTime.TotalMilliseconds - $t0[$p.Id]) / (20.0 * $cores) } } catch { }
                   New-Object PSObject -Property @{ Name = $p.ProcessName; Id = $p.Id; Cpu = $cpu; Mem = [double]$p.WorkingSet64 }
               }
               Head 'Procesor (% całego CPU)'
               $rows | Sort-Object Cpu -Descending | Select-Object -First 10 | ForEach-Object { Log ('{0,6:N1}%   {1} (PID {2})' -f $_.Cpu, $_.Name, $_.Id) 'out' }
               Head 'Pamięć (zestaw roboczy)'
               $rows | Sort-Object Mem -Descending | Select-Object -First 10 | ForEach-Object { Log ('{0,10}   {1} (PID {2})' -f (Format-Size $_.Mem), $_.Name, $_.Id) 'out' }
           } }
        @{ Cat = 'mem'; Title = 'Plik stronicowania: automatyczny'; Tags = @('restart')
           Desc = 'Naprawia błędy "za mało pamięci" po ręcznym ustawieniu zbyt małego pliku.'
           Code = {
               $cs = Get-CimInstance Win32_ComputerSystem
               if ($cs.AutomaticManagedPagefile) { Ok 'Plik stronicowania już jest zarządzany automatycznie.'; return }
               try { Set-CimInstance -InputObject $cs -Property @{ AutomaticManagedPagefile = $true } -ErrorAction Stop; Ok 'Ustawiono automatyczne zarządzanie plikiem stronicowania – wymagany restart.' }
               catch { Err $_.Exception.Message }
           } }
        @{ Cat = 'mem'; Title = 'Test pamięci RAM'; Tags = @('restart')
           Desc = 'Uruchamia Diagnostykę pamięci systemu Windows (test wykonywany przy restarcie).'
           Ui = { try { Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\mdsched.exe') } catch { Show-Error $_.Exception.Message } } }

        # ------------------------------------------------------------ PERFORMANCE
        @{ Cat = 'perf'; Title = 'Plan zasilania: maksymalna wydajność'; Tags = @('PC stacjonarny')
           Desc = 'Włącza plan "Najwyższa wydajność" (na laptopach skraca czas pracy na baterii).'
           Code = {
               $pc = Join-Path $env:SystemRoot 'System32\powercfg.exe'
               $myGuid = '5ce4e2a1-8d1e-4b9f-9a77-2b7d0c0f3a11'
               $list = (& $pc /list) -join "`n"
               if ($list -notmatch $myGuid) { [void](Invoke-Cmd $pc ('/duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 ' + $myGuid)) }
               $c = Invoke-Cmd $pc ('/setactive ' + $myGuid)
               if ($c -ne 0) { $c = Invoke-Cmd $pc '/setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c' }
               if ($c -eq 0) { Ok 'Aktywny plan wysokiej wydajności.' }
               else { Warn 'Ten komputer (Modern Standby) obsługuje tylko plan Zrównoważony – użyj suwaka "Tryb zasilania" w Ustawieniach > System > Zasilanie.' }
           } }
        @{ Cat = 'perf'; Title = 'Plan zasilania: zrównoważony'; Tags = @('domyślny')
           Desc = 'Przywraca domyślny plan zasilania Windows.'
           Code = { $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\powercfg.exe') '/setactive 381b4222-f694-41f0-9685-ff5bb260df2e'; if ($c -eq 0) { Ok 'Aktywny plan Zrównoważony.' } } }
        @{ Cat = 'perf'; Title = 'Przyspiesz interfejs'; Tags = @('wylogowanie')
           Desc = 'Wyłącza animacje, cienie i przezroczystość, skraca opóźnienie menu i autostartu.'
           Code = {
               Backup-RegKey 'HKEY_CURRENT_USER\Control Panel\Desktop' 'Desktop'
               Backup-RegKey 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ExplorerAdvanced'
               Backup-RegKey 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualEffects'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualFXSetting' 3 'DWord'
               Set-Reg 'HKCU:\Control Panel\Desktop' 'UserPreferencesMask' ([byte[]](0x90, 0x12, 0x03, 0x80, 0x10, 0x00, 0x00, 0x00)) 'Binary'
               Set-Reg 'HKCU:\Control Panel\Desktop' 'MenuShowDelay' '100' 'String'
               Set-Reg 'HKCU:\Control Panel\Desktop\WindowMetrics' 'MinAnimate' '0' 'String'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'TaskbarAnimations' 0 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ListviewAlphaSelect' 0 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ListviewShadow' 0 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\DWM' 'EnableAeroPeek' 0 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'EnableTransparency' 0 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize' 'StartupDelayInMSec' 0 'DWord'
               Ok 'Ustawienia zapisane – wyloguj się i zaloguj ponownie, aby zadziałały w pełni.'
           } }
        @{ Cat = 'perf'; Title = 'Przywróć domyślny wygląd'; Tags = @('wylogowanie')
           Desc = 'Cofa zmiany "Przyspiesz interfejs": animacje, przezroczystość, opóźnienia.'
           Code = {
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects' 'VisualFXSetting' 0 'DWord'
               Set-Reg 'HKCU:\Control Panel\Desktop' 'UserPreferencesMask' ([byte[]](0x9E, 0x1E, 0x07, 0x80, 0x12, 0x00, 0x00, 0x00)) 'Binary'
               Set-Reg 'HKCU:\Control Panel\Desktop' 'MenuShowDelay' '400' 'String'
               Set-Reg 'HKCU:\Control Panel\Desktop\WindowMetrics' 'MinAnimate' '1' 'String'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'TaskbarAnimations' 1 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ListviewAlphaSelect' 1 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' 'ListviewShadow' 1 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\DWM' 'EnableAeroPeek' 1 'DWord'
               Set-Reg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'EnableTransparency' 1 'DWord'
               Remove-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize' -Name 'StartupDelayInMSec' -ErrorAction SilentlyContinue
               Ok 'Przywrócono domyślne ustawienia – wyloguj się i zaloguj ponownie.'
           } }
        @{ Cat = 'perf'; Title = 'Raport czasu uruchamiania'; Tags = @('diagnostyka')
           Desc = 'Ile trwa start Windows i które programy, usługi lub sterowniki go spowalniają.'
           Code = {
               $log = 'Microsoft-Windows-Diagnostics-Performance/Operational'
               function Get-EvData($e) { $h = @{}; try { foreach ($d in ([xml]$e.ToXml()).Event.EventData.Data) { $h[$d.Name] = $d.'#text' } } catch { }; return $h }
               Head 'Ostatnie uruchomienia'
               $ev = @(Get-WinEvent -FilterHashtable @{ LogName = $log; Id = 100 } -MaxEvents 8 -ErrorAction SilentlyContinue)
               if ($ev.Count -eq 0) { Warn 'Brak danych (dziennik diagnostyki wydajności jest pusty lub wyłączony).' }
               foreach ($e in $ev) {
                   $d = Get-EvData $e
                   Log ('{0:yyyy-MM-dd HH:mm}   start: {1,5:N1} s   (do pulpitu {2:N1} s, po zalogowaniu {3:N1} s)' -f $e.TimeCreated, ([double]$d['BootTime'] / 1000), ([double]$d['MainPathBootTime'] / 1000), ([double]$d['BootPostBootTime'] / 1000)) 'out'
               }
               Head 'Co spowalnia start (30 dni)'
               $slow = @(Get-WinEvent -FilterHashtable @{ LogName = $log; Id = 101, 102, 103, 106, 109; StartTime = (Get-Date).AddDays(-30) } -ErrorAction SilentlyContinue)
               if ($slow.Count -eq 0) { Ok 'Nie odnotowano elementów spowalniających start.'; return }
               $kinds = @{ 101 = 'aplikacja'; 102 = 'sterownik'; 103 = 'usługa'; 106 = 'zadanie w tle'; 109 = 'urządzenie' }
               $slow | ForEach-Object {
                   $d = Get-EvData $_
                   $n = $d['FriendlyName']; if (-not $n) { $n = $d['Name'] }
                   New-Object PSObject -Property @{ Name = ('{0} [{1}]' -f $n, $kinds[$_.Id]); Deg = [double]$d['DegradationTime'] }
               } | Group-Object Name | ForEach-Object {
                   New-Object PSObject -Property @{ Name = $_.Name; Count = $_.Count; Avg = (($_.Group | Measure-Object Deg -Average).Average) }
               } | Sort-Object Avg -Descending | Select-Object -First 12 | ForEach-Object {
                   Log ('+{0,6:N1} s średnio  ({1}x)   {2}' -f ($_.Avg / 1000), $_.Count, $_.Name) 'warn'
               }
               Log 'Wskazówka: zbędne programy wyłączysz w zakładce Autostart.'
           } }
        @{ Cat = 'perf'; Title = 'Odbuduj liczniki wydajności'; Tags = @()
           Desc = 'Gdy Menedżer zadań lub Monitor wydajności pokazuje błędy albo puste wykresy.'
           Code = {
               $s32 = Join-Path $env:SystemRoot 'System32'
               [void](Invoke-Cmd (Join-Path $s32 'lodctr.exe') '/R' -WorkDir $s32)
               $wow = Join-Path $env:SystemRoot 'SysWOW64'
               if (Test-Path -LiteralPath (Join-Path $wow 'lodctr.exe')) { [void](Invoke-Cmd (Join-Path $wow 'lodctr.exe') '/R' -WorkDir $wow) }
               [void](Invoke-Cmd (Join-Path $s32 'wbem\WinMgmt.exe') '/resyncperf')
               Ok 'Liczniki wydajności odbudowane.'
           } }
        @{ Cat = 'perf'; Title = 'Menedżer autostartu'; Tags = @('przejdź')
           Desc = 'Włączaj i wyłączaj programy uruchamiane razem z Windows.'
           Ui = { $ui.tabMain.SelectedItem = $ui.tabStart; Update-StartupList } }

        # ------------------------------------------------------------ EXPLORER / UI
        @{ Cat = 'ui'; Title = 'Uruchom ponownie Eksplorator'; Tags = @('szybkie')
           Desc = 'Gdy zawiesił się pasek zadań, pulpit lub okna folderów.'
           Code = { Restart-Explorer } }
        @{ Cat = 'ui'; Title = 'Napraw menu Start i pasek zadań'; Tags = @()
           Desc = 'Ponownie rejestruje składniki Start, wyszukiwania i paska zadań.'
           Code = {
               $names = @('Microsoft.Windows.StartMenuExperienceHost', 'Microsoft.Windows.ShellExperienceHost', 'MicrosoftWindows.Client.CBS', 'MicrosoftWindows.Client.Core', 'Microsoft.Windows.Search')
               $seen = @{}
               foreach ($n in $names) {
                   foreach ($p in @(Get-AppxPackage -AllUsers -Name $n -ErrorAction SilentlyContinue)) {
                       if ($seen.ContainsKey($p.PackageFullName)) { continue }
                       $seen[$p.PackageFullName] = 1
                       $m = Join-Path $p.InstallLocation 'AppxManifest.xml'
                       if (-not (Test-Path -LiteralPath $m)) { continue }
                       try { Add-AppxPackage -DisableDevelopmentMode -Register $m -ErrorAction Stop; Ok ('Zarejestrowano: {0}' -f $p.Name) }
                       catch { Warn ('{0}: {1}' -f $p.Name, $_.Exception.Message) }
                   }
               }
               foreach ($pn in @('StartMenuExperienceHost', 'SearchHost', 'ShellExperienceHost')) { Get-Process -Name $pn -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }
               Restart-Explorer
           } }
        @{ Cat = 'ui'; Title = 'Przebuduj ikony i miniatury'; Tags = @()
           Desc = 'Naprawia puste, błędne lub nieaktualne ikony i podglądy plików.'
           Code = {
               $dir = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Explorer'
               $ok = 0; $fail = 0
               for ($round = 0; $round -lt 3; $round++) {
                   Get-SessionExplorer | Stop-Process -Force -ErrorAction SilentlyContinue
                   Start-Sleep -Milliseconds 800
                   $files = @(Get-ChildItem -LiteralPath $dir -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'iconcache*.db' -or $_.Name -like 'thumbcache*.db' })
                   $legacy = Join-Path $env:LOCALAPPDATA 'IconCache.db'
                   if (Test-Path -LiteralPath $legacy) { $files += Get-Item -LiteralPath $legacy -Force }
                   $fail = 0
                   foreach ($f in $files) { try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $ok++ } catch { $fail++ } }
                   if ($fail -eq 0) { break }
               }
               if ((Get-SessionExplorer).Count -eq 0) { Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') }
               try { Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\ie4uinit.exe') -ArgumentList '-show' -WindowStyle Hidden } catch { }
               [RegDoktor.Native]::RefreshShellAssociations()
               Ok ('Usunięto plików pamięci podręcznej: {0}{1}' -f $ok, $(if ($fail) { ' (zablokowanych: ' + $fail + ' – usuną się po restarcie)' } else { '' }))
           } }
        @{ Cat = 'ui'; Title = 'Klasyczne menu kontekstowe'; Tags = @('przełącznik')
           Desc = 'Pełne menu prawego przycisku myszy bez "Pokaż więcej opcji" (lub powrót do nowego).'
           Confirm = {
               if (Test-Path 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}') { 'Włączone jest KLASYCZNE menu kontekstowe. Przywrócić nowe menu Windows 11? (Eksplorator uruchomi się ponownie)' }
               else { 'Włączyć KLASYCZNE (pełne) menu kontekstowe? (Eksplorator uruchomi się ponownie)' }
           }
           Code = {
               $k = 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
               if (Test-Path -LiteralPath $k) { Remove-Item -LiteralPath $k -Recurse -Force; Ok 'Przywrócono nowe menu kontekstowe Windows 11.' }
               else {
                   New-Item -Path ($k + '\InprocServer32') -Force | Out-Null
                   [RegDoktor.Reg]::SetValue('HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32', '', [Microsoft.Win32.RegistryValueKind]::String, '')
                   Ok 'Włączono klasyczne menu kontekstowe.'
               }
               Restart-Explorer
           } }
        @{ Cat = 'ui'; Title = 'Pokaż rozszerzenia i ukryte pliki'; Tags = @('przełącznik')
           Desc = 'Widoczne rozszerzenia (.exe, .pdf…) chronią przed fałszywymi plikami.'
           Confirm = {
               $v = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -ErrorAction SilentlyContinue).HideFileExt
               if ($v -eq 0) { 'Rozszerzenia są teraz WIDOCZNE. Ukryć rozszerzenia i ukryte pliki (ustawienie domyślne)?' } else { 'Pokazać rozszerzenia plików i ukryte pliki?' }
           }
           Code = {
               $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
               if ((Get-RegValue $k 'HideFileExt') -eq 0) { Set-Reg $k 'HideFileExt' 1 'DWord'; Set-Reg $k 'Hidden' 2 'DWord'; Ok 'Rozszerzenia i ukryte pliki są teraz ukryte.' }
               else { Set-Reg $k 'HideFileExt' 0 'DWord'; Set-Reg $k 'Hidden' 1 'DWord'; Ok 'Rozszerzenia i ukryte pliki są teraz widoczne.' }
               [RegDoktor.Native]::RefreshShellAssociations()
           } }
        @{ Cat = 'ui'; Title = 'Napraw Szybki dostęp'; Tags = @()
           Desc = 'Gdy Eksplorator zawiesza się na "Szybkim dostępie". Resetuje przypięte foldery.'
           Confirm = 'Przypięte foldery w Szybkim dostępie zostaną zresetowane do domyślnych. Kontynuować?'
           Code = {
               $d = Join-Path $env:APPDATA 'Microsoft\Windows\Recent\AutomaticDestinations'
               Get-SessionExplorer | Stop-Process -Force -ErrorAction SilentlyContinue
               Start-Sleep -Milliseconds 800
               foreach ($f in @('f01b4d95cf55d32a.automaticDestinations-ms', '5f7b5f1e01b83767.automaticDestinations-ms')) {
                   $p = Join-Path $d $f
                   if (Test-Path -LiteralPath $p) { try { Remove-Item -LiteralPath $p -Force -ErrorAction Stop; Log ('Usunięto ' + $f) 'out' } catch { Warn $_.Exception.Message } }
               }
               if ((Get-SessionExplorer).Count -eq 0) { Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') }
               Ok 'Szybki dostęp zresetowany.'
           } }
        @{ Cat = 'ui'; Title = 'Napraw skojarzenia .exe / .lnk'; Tags = @(); Risky = $true
           Desc = 'Gdy programy lub skróty otwierają się w złej aplikacji (np. po wirusie).'
           Confirm = 'Przywrócić domyślne skojarzenia plików .exe, .lnk, .bat, .cmd, .com i .reg? (Eksplorator uruchomi się ponownie)'
           Code = {
               $map = [ordered]@{ '.exe' = 'exefile'; '.lnk' = 'lnkfile'; '.bat' = 'batfile'; '.cmd' = 'cmdfile'; '.com' = 'comfile'; '.reg' = 'regfile' }
               foreach ($ext in $map.Keys) {
                   Backup-RegKey ('HKEY_LOCAL_MACHINE\SOFTWARE\Classes\' + $ext) ('HKLM_' + $ext.TrimStart('.'))
                   Backup-RegKey ('HKEY_CURRENT_USER\Software\Classes\' + $ext) ('HKCU_' + $ext.TrimStart('.'))
                   Backup-RegKey ('HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\' + $ext) ('FileExts_' + $ext.TrimStart('.'))
                   [RegDoktor.Reg]::SetValue(('HKEY_LOCAL_MACHINE\SOFTWARE\Classes\' + $ext), '', [Microsoft.Win32.RegistryValueKind]::String, $map[$ext])
                   if ([RegDoktor.Reg]::KeyExists('HKEY_CURRENT_USER\Software\Classes\' + $ext)) { [RegDoktor.Reg]::DeleteKeyTree('HKEY_CURRENT_USER\Software\Classes\' + $ext) }
                   $uc = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\' + $ext + '\UserChoice'
                   if ([RegDoktor.Reg]::KeyExists($uc)) { try { [RegDoktor.Reg]::DeleteKeyTree($uc) } catch { Warn ('{0}: {1}' -f $uc, $_.Exception.Message) } }
                   Ok ('{0} -> {1}' -f $ext, $map[$ext])
               }
               [RegDoktor.Native]::RefreshShellAssociations()
               Restart-Explorer
           } }

        # ------------------------------------------------------------ DEVICES
        @{ Cat = 'dev'; Title = 'Napraw kolejkę drukowania'; Tags = @('zalecane')
           Desc = 'Gdy wydruk "wisi" w kolejce i nic się nie drukuje.'
           Code = {
               Stop-ServiceSafe 'Spooler'
               $d = Join-Path $env:SystemRoot 'System32\spool\PRINTERS'
               $n = 0
               foreach ($f in @(Get-ChildItem -LiteralPath $d -Force -ErrorAction SilentlyContinue)) { try { Remove-Item -LiteralPath $f.FullName -Force -Recurse -ErrorAction Stop; $n++ } catch { } }
               Start-ServiceSafe 'Spooler'
               Ok ('Usunięto zadań z kolejki: {0}. Bufor wydruku działa.' -f $n)
           } }
        @{ Cat = 'dev'; Title = 'Napraw dźwięk'; Tags = @('szybkie')
           Desc = 'Uruchamia ponownie usługi audio – gdy nagle zniknął dźwięk.'
           Code = {
               try { Restart-Service -Name 'AudioEndpointBuilder' -Force -ErrorAction Stop; Log 'AudioEndpointBuilder uruchomiony ponownie' } catch { Warn $_.Exception.Message }
               Start-ServiceSafe 'Audiosrv'
               Ok 'Usługi audio działają. Jeśli nadal brak dźwięku – sprawdź urządzenie wyjściowe w Ustawieniach dźwięku.'
           } }
        @{ Cat = 'dev'; Title = 'Urządzenia z problemami'; Tags = @('diagnostyka')
           Desc = 'Lista urządzeń z błędami sterowników (jak żółte wykrzykniki w Menedżerze urządzeń).'
           Code = {
               $codes = @{ 1 = 'nieskonfigurowane'; 3 = 'uszkodzony sterownik / brak pamięci'; 10 = 'nie można uruchomić'; 12 = 'konflikt zasobów'; 14 = 'wymaga restartu'; 18 = 'zainstaluj ponownie sterownik'; 19 = 'błąd w rejestrze'; 21 = 'usuwane'; 22 = 'wyłączone'; 24 = 'brak lub źle zainstalowane'; 28 = 'brak sterownika'; 31 = 'sterownik nie działa'; 32 = 'sterownik wyłączony'; 37 = 'błąd inicjalizacji sterownika'; 39 = 'uszkodzony/brakujący sterownik'; 43 = 'urządzenie zgłosiło błąd'; 48 = 'sterownik zablokowany'; 52 = 'brak podpisu sterownika' }
               $bad = @(Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue | Where-Object { $_.ConfigManagerErrorCode -ne 0 -and $_.ConfigManagerErrorCode -ne 45 })
               if ($bad.Count -eq 0) { Ok 'Wszystkie podłączone urządzenia działają poprawnie.'; return }
               foreach ($b in $bad) {
                   $c = [int]$b.ConfigManagerErrorCode
                   $txt = $codes[$c]; if (-not $txt) { $txt = 'błąd' }
                   Warn ('{0}  [{1}]  kod {2}: {3}' -f $b.Name, $b.PNPClass, $c, $txt)
               }
               Log 'Wskazówka: "Skanuj zmiany sprzętu" lub aktualizacja sterownika od producenta zwykle pomaga.'
           } }
        @{ Cat = 'dev'; Title = 'Skanuj zmiany sprzętu'; Tags = @()
           Desc = 'Ponownie wykrywa urządzenia i instaluje brakujące sterowniki.'
           Code = { $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\pnputil.exe') '/scan-devices'; if ($c -eq 0) { Ok 'Skanowanie zakończone.' } } }
        @{ Cat = 'dev'; Title = 'Napraw Bluetooth'; Tags = @()
           Desc = 'Uruchamia ponownie usługi Bluetooth (gdy urządzenia nie łączą się).'
           Code = {
               foreach ($s in @(Get-Service -Name 'bthserv', 'BthAvctpSvc', 'BluetoothUserService_*' -ErrorAction SilentlyContinue)) {
                   try { Restart-Service -Name $s.Name -Force -ErrorAction Stop; Ok ('Uruchomiono ponownie: {0}' -f $s.Name) } catch { Warn ('{0}: {1}' -f $s.Name, $_.Exception.Message) }
               }
           } }
        @{ Cat = 'dev'; Title = 'Raport baterii'; Tags = @('laptop', 'raport')
           Desc = 'Pojemność fabryczna vs. obecna, zużycie i historia ładowania (HTML).'
           Code = {
               $f = Join-Path (Get-ReportDir) 'raport-baterii.html'
               $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\powercfg.exe') ('/batteryreport /output "{0}"' -f $f)
               if ($c -eq 0 -and (Test-Path -LiteralPath $f)) { Open-Unelevated $f; Ok ('Raport: {0}' -f $f) } else { Warn 'Nie wygenerowano raportu (komputer bez baterii?).' }
           } }

        # ------------------------------------------------------------ SECURITY
        @{ Cat = 'sec'; Title = 'Stan zabezpieczeń'; Tags = @('diagnostyka')
           Desc = 'Antywirus, ochrona w czasie rzeczywistym, zapora, UAC, Secure Boot, TPM, BitLocker.'
           Code = {
               try {
                   foreach ($av in @(Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop)) { Log ('Antywirus: {0}' -f $av.displayName) }
               } catch { }
               $s = Get-MpComputerStatus -ErrorAction SilentlyContinue
               if ($s) {
                   if ($s.RealTimeProtectionEnabled) { Ok 'Defender: ochrona w czasie rzeczywistym włączona' } else { Warn 'Defender: ochrona w czasie rzeczywistym wyłączona (lub działa inny antywirus)' }
                   $age = [int]$s.AntivirusSignatureAge
                   if ($age -le 2) { Ok ('Definicje wirusów: aktualne ({0})' -f $s.AntivirusSignatureLastUpdated) } else { Warn ('Definicje wirusów mają {0} dni – zaktualizuj.' -f $age) }
                   Log ('Ostatnie szybkie skanowanie: {0} dni temu' -f $s.QuickScanAge) 'out'
               }
               foreach ($p in @(Get-NetFirewallProfile -ErrorAction SilentlyContinue)) {
                   if ($p.Enabled) { Ok ('Zapora ({0}): włączona' -f $p.Name) } else { Warn ('Zapora ({0}): WYŁĄCZONA' -f $p.Name) }
               }
               if ((Get-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' 'EnableLUA') -eq 0) { Warn 'Kontrola konta użytkownika (UAC): WYŁĄCZONA' } else { Ok 'Kontrola konta użytkownika (UAC): włączona' }
               try { if (Confirm-SecureBootUEFI -ErrorAction Stop) { Ok 'Secure Boot: włączony' } else { Warn 'Secure Boot: wyłączony' } } catch { Log 'Secure Boot: nieobsługiwany (BIOS legacy)' 'out' }
               try { $t = Get-Tpm -ErrorAction Stop; if ($t.TpmPresent -and $t.TpmReady) { Ok 'TPM: obecny i gotowy' } elseif ($t.TpmPresent) { Warn 'TPM: obecny, ale nie gotowy' } else { Warn 'TPM: brak' } } catch { }
               try { $b = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop; Log ('BitLocker {0}: {1}' -f $env:SystemDrive, $b.ProtectionStatus) 'out' } catch { }
           } }
        @{ Cat = 'sec'; Title = 'Aktualizuj definicje Defendera'; Tags = @('internet')
           Desc = 'Pobiera najnowsze sygnatury zagrożeń Microsoft Defender.'
           Code = { try { Update-MpSignature -ErrorAction Stop; Ok 'Definicje zaktualizowane.' } catch { Err $_.Exception.Message } } }
        @{ Cat = 'sec'; Title = 'Szybkie skanowanie Defenderem'; Tags = @('5–15 min')
           Desc = 'Aktualizuje definicje i skanuje miejsca najczęściej atakowane.'
           Code = {
               try { Update-MpSignature -ErrorAction Stop; Log 'Definicje zaktualizowane.' } catch { Warn ('Aktualizacja definicji: ' + $_.Exception.Message) }
               Progress -1 'Skanowanie Microsoft Defender…'
               try { Start-MpScan -ScanType QuickScan -ErrorAction Stop; Ok 'Skanowanie zakończone.' } catch { Err $_.Exception.Message; return }
               $th = @(Get-MpThreatDetection -ErrorAction SilentlyContinue | Where-Object { $_.InitialDetectionTime -gt (Get-Date).AddHours(-1) })
               if ($th.Count) { foreach ($t in $th) { Warn ('Wykryto: {0}' -f ($t.Resources -join ', ')) } } else { Ok 'Nie wykryto zagrożeń.' }
           } }
        @{ Cat = 'sec'; Title = 'Skanowanie offline Defenderem'; Tags = @('restart natychmiast')
           Desc = 'Uruchamia komputer ponownie i skanuje go przed startem Windows (uparty malware).'
           Confirm = 'Komputer ZARAZ uruchomi się ponownie i przez ok. 15 minut będzie skanowany. Zapisz otwarte pliki! Kontynuować?'
           Code = { try { Start-MpWDOScan -ErrorAction Stop; Ok 'Skanowanie offline zaplanowane – trwa restart.' } catch { Err $_.Exception.Message } } }
        @{ Cat = 'sec'; Title = 'Przywróć domyślną zaporę'; Tags = @(); Risky = $true
           Desc = 'Usuwa własne reguły zapory i przywraca ustawienia fabryczne.'
           Confirm = 'Wszystkie własne reguły zapory Windows zostaną usunięte (niektóre programy zapytają ponownie o dostęp do sieci). Kontynuować?'
           Code = { $c = Invoke-Cmd (Join-Path $env:SystemRoot 'System32\netsh.exe') 'advfirewall reset'; if ($c -eq 0) { Ok 'Zapora przywrócona do ustawień domyślnych.' } } }

        # ------------------------------------------------------------ REGISTRY
        @{ Cat = 'reg'; Title = 'Pełna kopia rejestru'; Tags = @('zalecane')
           Desc = 'Zapisuje gałęzie SOFTWARE, SYSTEM, DEFAULT i Twój profil (NTUSER, UsrClass).'
           Code = {
               $dir = Join-Path $BackupRoot ('Rejestr_' + (Get-Stamp))
               New-Item -ItemType Directory -Path $dir -Force | Out-Null
               $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
               $list = @(@('HKLM', 'SOFTWARE', 'SOFTWARE.hiv'), @('HKLM', 'SYSTEM', 'SYSTEM.hiv'), @('HKU', '.DEFAULT', 'DEFAULT.hiv'), @('HKU', $sid, 'NTUSER.hiv'), @('HKU', ($sid + '_Classes'), 'UsrClass.hiv'))
               $i = 0
               foreach ($h in $list) {
                   $i++; Progress ($i * 100 / $list.Count) ('Zapis gałęzi ' + $h[1])
                   $f = Join-Path $dir $h[2]
                   $rc = [RegDoktor.Native]::SaveHive($h[0], $h[1], $f)
                   if ($rc -eq 0) { Ok ('{0}\{1} -> {2} ({3})' -f $h[0], $h[1], $h[2], (Format-Size (Get-Item -LiteralPath $f).Length)) }
                   else { Err ('{0}\{1}: {2}' -f $h[0], $h[1], (New-Object System.ComponentModel.Win32Exception $rc).Message) }
               }
               $txt = @"
Kopia rejestru wykonana przez RegDoktor 11: $(Get-Date)
Konto: $env:USERDOMAIN\$env:USERNAME  (SID: $sid)

Pliki .hiv to binarne kopie gałęzi rejestru. Przywraca się je, gdy Windows nie działa:
 1. Uruchom Środowisko odzyskiwania (Shift + Uruchom ponownie) > Rozwiązywanie problemów > Wiersz polecenia.
 2. Sprawdź literę dysku z Windows (np. dir C:\ , dir D:\ ).
 3. Skopiuj pliki, np.:
      copy /y "<ta_kopia>\SOFTWARE.hiv" C:\Windows\System32\config\SOFTWARE
      copy /y "<ta_kopia>\SYSTEM.hiv"   C:\Windows\System32\config\SYSTEM
      copy /y "<ta_kopia>\DEFAULT.hiv"  C:\Windows\System32\config\DEFAULT
      copy /y "<ta_kopia>\NTUSER.hiv"   C:\Users\<Twoje_konto>\NTUSER.DAT
      copy /y "<ta_kopia>\UsrClass.hiv" C:\Users\<Twoje_konto>\AppData\Local\Microsoft\Windows\UsrClass.dat
Pojedyncze klucze możesz też podejrzeć: regedit > Plik > Załaduj gałąź.
"@
               Set-Content -LiteralPath (Join-Path $dir 'PRZYWRACANIE.txt') -Value $txt -Encoding UTF8
               Ok ('Kopia zapisana w: {0}' -f $dir)
           } }
        @{ Cat = 'reg'; Title = 'Skanuj osierocone wpisy'; Tags = @('przejdź')
           Desc = 'Wpisy po odinstalowanych programach, brakujące pliki, puste klucze.'
           Ui = { $ui.tabMain.SelectedItem = $ui.tabClean; Start-OrphanScan } }
        @{ Cat = 'reg'; Title = 'Defragmentuj rejestr'; Tags = @('przejdź')
           Desc = 'Analiza i kompaktowanie plików gałęzi rejestru.'
           Ui = { $ui.tabMain.SelectedItem = $ui.tabHive; Start-HiveAnalyze } }
        @{ Cat = 'reg'; Title = 'Wyczyść historię (MRU)'; Tags = @('prywatność')
           Desc = 'Historia okna Uruchom, paska adresu, wyszukiwania i ostatnio otwieranych plików.'
           Confirm = 'Usunąć historię okna Uruchom, wpisanych ścieżek, wyszukiwań w Eksploratorze i listy ostatnich plików? (kopia .reg zostanie zapisana)'
           Code = {
               $base = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\'
               foreach ($k in @('RunMRU', 'TypedPaths', 'RecentDocs', 'WordWheelQuery', 'ComDlg32\OpenSavePidlMRU', 'ComDlg32\LastVisitedPidlMRU')) {
                   $p = $base + $k
                   if (-not [RegDoktor.Reg]::KeyExists($p)) { continue }
                   Backup-RegKey $p ('MRU_' + ($k -replace '\\', '_'))
                   try { [RegDoktor.Reg]::DeleteKeyTree($p); Ok ('Wyczyszczono: ' + $k) } catch { Warn ('{0}: {1}' -f $k, $_.Exception.Message) }
               }
               $recent = Join-Path $env:APPDATA 'Microsoft\Windows\Recent'
               $n = 0
               foreach ($f in @(Get-ChildItem -LiteralPath $recent -Filter *.lnk -Force -ErrorAction SilentlyContinue)) { try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $n++ } catch { } }
               Ok ('Usunięto skrótów ostatnich plików: {0}' -f $n)
           } }
        @{ Cat = 'reg'; Title = 'Przywróć z kopii .reg'; Tags = @()
           Desc = 'Importuje wybrane pliki kopii zapasowych utworzonych przez RegDoktor.'
           Ui = { Restore-RegBackup } }
    )

    # =====================================================================
    #  ORPHAN SCAN / CLEAN, HIVE ANALYSIS / COMPACTION (background jobs)
    # =====================================================================
    $ScanCategories = @(
        @{ Id = 'uninstall';  Risk = 1; Name = 'Odinstalowane programy';        Desc = 'Wpisy w "Programy i funkcje", których deinstalator i folder już nie istnieją.' }
        @{ Id = 'apppaths';   Risk = 1; Name = 'Ścieżki aplikacji (App Paths)'; Desc = 'Skróty uruchamiania programów wskazujące na usunięte pliki .exe.' }
        @{ Id = 'run';        Risk = 1; Name = 'Autostart bez pliku';           Desc = 'Wpisy Run/RunOnce uruchamiające nieistniejące programy.' }
        @{ Id = 'shareddlls'; Risk = 1; Name = 'Współdzielone biblioteki';      Desc = 'Liczniki SharedDLLs dla bibliotek, których już nie ma.' }
        @{ Id = 'muicache';   Risk = 1; Name = 'Pamięć MUI (MuiCache)';         Desc = 'Zapamiętane nazwy programów, które zostały usunięte.' }
        @{ Id = 'folders';    Risk = 1; Name = 'Foldery instalatora';           Desc = 'Rejestr folderów Instalatora Windows, które nie istnieją.' }
        @{ Id = 'apps';       Risk = 1; Name = 'Aplikacje "Otwórz za pomocą"';  Desc = 'Programy na liście "Otwórz za pomocą", których pliki usunięto.' }
        @{ Id = 'fonts';      Risk = 1; Name = 'Czcionki bez plików';           Desc = 'Zarejestrowane czcionki, których pliki zostały usunięte.' }
        @{ Id = 'msi';        Risk = 2; Name = 'Osierocone wpisy MSI';          Desc = 'Wpisy deinstalacji MSI bez rejestracji w Instalatorze Windows.' }
        @{ Id = 'approved';   Risk = 2; Name = 'Pozostałości StartupApproved';  Desc = 'Stany włącz/wyłącz autostartu dla programów, których już nie ma.' }
        @{ Id = 'com';        Risk = 2; Name = 'Klasy COM / ActiveX bez plików'; Desc = 'Rejestracje CLSID wskazujące na usunięte biblioteki i programy.' }
        @{ Id = 'empty';      Risk = 2; Name = 'Puste klucze (węzły)';          Desc = 'Klucze oprogramowania bez żadnych wartości i podkluczy.' }
        @{ Id = 'services';   Risk = 3; Name = 'Usługi i sterowniki bez plików'; Desc = 'Wpisy usług wskazujące na brakujące pliki. Tylko dla zaawansowanych!' }
    )

    $OrphanScanCode = {
        $found = New-Object 'System.Collections.Generic.List[RegDoktor.CleanItem]'
        $riskNames = @('', 'Niskie', 'Średnie', 'Wysokie')
        function Add-Hit([string]$Category, [int]$Risk, [string]$Key, $Value, [bool]$WholeKey, [string]$Target, [string]$Detail) {
            $o = New-Object RegDoktor.CleanItem
            $o.Category = $Category; $o.RiskLevel = $Risk; $o.Risk = $riskNames[$Risk]
            $o.KeyPath = [RegDoktor.Reg]::NormalizePath($Key)
            if ($WholeKey) { $o.ValueName = $null } else { $o.ValueName = [string]$Value }
            $o.IsWholeKey = $WholeKey; $o.Target = $Target; $o.Detail = $Detail
            $o.Selected = ($Risk -eq 1)
            $found.Add($o)
        }
        function Open-Key([string]$Path) { return [RegDoktor.Reg]::TryOpen($Path, $false) }
        function Get-Str($Key, [string]$Name) { try { return [string]$Key.GetValue($Name, '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) } catch { return '' } }
        $script:MsiSids = $null
        function Test-MsiRegistered([string]$Code) {
            $p = ConvertTo-PackedGuid $Code
            if (-not $p) { return $true }
            if ([RegDoktor.Reg]::KeyExists('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed')) { return $true }
            if ([RegDoktor.Reg]::KeyExists('HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Installer\Products\' + $p)) { return $true }
            if ([RegDoktor.Reg]::KeyExists('HKEY_CURRENT_USER\Software\Microsoft\Installer\Products\' + $p)) { return $true }
            $ud = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData'
            if ($null -eq $script:MsiSids) {
                $script:MsiSids = @()
                $k = Open-Key $ud
                if ($k) { $script:MsiSids = $k.GetSubKeyNames(); $k.Close() }
            }
            foreach ($s in $script:MsiSids) { if ([RegDoktor.Reg]::KeyExists($ud + '\' + $s + '\Products\' + $p)) { return $true } }
            return $false
        }

        $steps = @($Cats).Count
        $step = 0
        function Step([string]$Text) { $script:step++; Progress ($script:step * 100.0 / ($steps + 1)) $Text; Log ('Skanowanie: ' + $Text) 'out' }

        # --- uninstall entries + MSI
        if ($Cats -contains 'uninstall' -or $Cats -contains 'msi') {
            Step 'odinstalowane programy'
            foreach ($r in @('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                             'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                             'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall')) {
                $k = Open-Key $r
                if (-not $k) { continue }
                try {
                    foreach ($n in $k.GetSubKeyNames()) {
                        $sk = $null
                        try { $sk = $k.OpenSubKey($n, $false) } catch { }
                        if (-not $sk) { continue }
                        try {
                            $dn = Get-Str $sk 'DisplayName'
                            $sc = $sk.GetValue('SystemComponent')
                            $wi = $sk.GetValue('WindowsInstaller')
                            $us = Get-Str $sk 'UninstallString'
                            $il = Get-Str $sk 'InstallLocation'
                            $pk = Get-Str $sk 'ParentKeyName'
                        } finally { $sk.Close() }
                        if (-not $dn -or $sc -eq 1 -or $pk) { continue }
                        if ($wi -eq 1 -or $us -match '(?i)msiexec') {
                            if ($Cats -contains 'msi' -and $n -match '^\{[0-9A-Fa-f-]{36}\}$' -and -not (Test-MsiRegistered $n)) {
                                Add-Hit 'Instalator MSI' 2 ($r + '\' + $n) $null $true '' ('"{0}" – produkt nie jest zarejestrowany w Instalatorze Windows' -f $dn)
                            }
                            continue
                        }
                        if ($Cats -notcontains 'uninstall' -or -not $us) { continue }
                        $miss = Get-MissingTarget $us
                        if (-not $miss) { continue }
                        if ($il) {
                            $ilp = Get-CmdPath ('"' + $il.Trim().Trim('"').TrimEnd('\') + '"')
                            if ($ilp -and (Test-PathExistsSmart $ilp)) { continue }
                        }
                        Add-Hit 'Odinstalowane programy' 1 ($r + '\' + $n) $null $true $miss ('"{0}" – brak deinstalatora: {1}' -f $dn, $miss)
                    }
                } finally { $k.Close() }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'apppaths') {
            Step 'ścieżki aplikacji'
            foreach ($r in @('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths',
                             'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths',
                             'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths')) {
                foreach ($pair in [RegDoktor.Reg]::CollectChildValues($r, '', '')) {
                    $miss = Get-MissingTarget $pair[1]
                    if ($miss) { Add-Hit 'Ścieżki aplikacji' 1 ($r + '\' + $pair[0]) $null $true $miss ('Brak pliku: ' + $miss) }
                }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'run') {
            Step 'autostart'
            foreach ($r in @('HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run',
                             'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce',
                             'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                             'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',
                             'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
                             'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce')) {
                $k = Open-Key $r
                if (-not $k) { continue }
                try {
                    foreach ($n in $k.GetValueNames()) {
                        if (-not $n) { continue }
                        $miss = Get-MissingTarget (Get-Str $k $n)
                        if ($miss) { Add-Hit 'Autostart' 1 $r $n $false $miss ('Brak pliku: ' + $miss) }
                    }
                } finally { $k.Close() }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'approved') {
            Step 'StartupApproved'
            $ap = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\'
            $pairs = @(
                @(('HKEY_CURRENT_USER\' + $ap + 'Run'), 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run', $null),
                @(('HKEY_LOCAL_MACHINE\' + $ap + 'Run'), 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run', $null),
                @(('HKEY_LOCAL_MACHINE\' + $ap + 'Run32'), 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run', $null),
                @(('HKEY_CURRENT_USER\' + $ap + 'StartupFolder'), $null, [Environment]::GetFolderPath('Startup')),
                @(('HKEY_LOCAL_MACHINE\' + $ap + 'StartupFolder'), $null, [Environment]::GetFolderPath('CommonStartup'))
            )
            foreach ($p in $pairs) {
                $k = Open-Key $p[0]
                if (-not $k) { continue }
                try {
                    foreach ($n in $k.GetValueNames()) {
                        if (-not $n) { continue }
                        $exists = $true
                        if ($p[1]) { $exists = [RegDoktor.Reg]::ValueExists($p[1], $n) }
                        elseif ($p[2]) { $exists = [System.IO.File]::Exists((Join-Path $p[2] $n)) }
                        if (-not $exists) { Add-Hit 'StartupApproved' 2 $p[0] $n $false '' ('Brak wpisu autostartu "{0}"' -f $n) }
                    }
                } finally { $k.Close() }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'shareddlls') {
            Step 'SharedDLLs'
            foreach ($r in @('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs', 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\SharedDLLs')) {
                $k = Open-Key $r
                if (-not $k) { continue }
                try {
                    foreach ($n in $k.GetValueNames()) {
                        $p = Get-CmdPath ('"' + $n + '"')
                        if ($p -and -not (Test-PathExistsSmart $p)) { Add-Hit 'SharedDLLs' 1 $r $n $false $p ('Brak pliku: ' + $p) }
                    }
                } finally { $k.Close() }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'muicache') {
            Step 'MuiCache'
            $r = 'HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache'
            $k = Open-Key $r
            if ($k) {
                try {
                    foreach ($n in $k.GetValueNames()) {
                        $m = [regex]::Match($n, '^([A-Za-z]:\\.+?)\.(FriendlyAppName|ApplicationCompany)$')
                        if (-not $m.Success) { continue }
                        $p = Get-CmdPath ('"' + $m.Groups[1].Value + '"')
                        if ($p -and -not (Test-PathExistsSmart $p)) { Add-Hit 'MuiCache' 1 $r $n $false $p ('Brak pliku: ' + $p) }
                    }
                } finally { $k.Close() }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'folders') {
            Step 'foldery instalatora'
            $r = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders'
            $k = Open-Key $r
            if ($k) {
                try {
                    foreach ($n in $k.GetValueNames()) {
                        $p = Get-CmdPath ('"' + $n.TrimEnd('\') + '"')
                        if ($p -and -not (Test-PathExistsSmart $p)) { Add-Hit 'Foldery instalatora' 1 $r $n $false $p ('Brak folderu: ' + $p) }
                    }
                } finally { $k.Close() }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'apps') {
            Step 'aplikacje "Otwórz za pomocą"'
            foreach ($r in @('HKEY_CURRENT_USER\Software\Classes\Applications', 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Applications')) {
                $risk = 1
                if ($r -like 'HKEY_LOCAL_MACHINE*') { $risk = 2 }
                foreach ($pair in [RegDoktor.Reg]::CollectChildValues($r, 'shell\open\command', '')) {
                    $miss = Get-MissingTarget $pair[1]
                    if ($miss) { Add-Hit 'Aplikacje' $risk ($r + '\' + $pair[0]) $null $true $miss ('Brak pliku: ' + $miss) }
                }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'fonts') {
            Step 'czcionki'
            $fontDir = Join-Path $env:SystemRoot 'Fonts'
            foreach ($r in @('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', 'HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Fonts')) {
                $k = Open-Key $r
                if (-not $k) { continue }
                try {
                    foreach ($n in $k.GetValueNames()) {
                        $v = [Environment]::ExpandEnvironmentVariables((Get-Str $k $n)).Trim()
                        if (-not $v) { continue }
                        try {
                            if ($v -match '^[A-Za-z]:\\') { $p = $v } else { $p = Join-Path $fontDir $v }
                            if ($p -match '^[A-Za-z]:\\' -and (Test-FixedDrive $p.Substring(0, 1)) -and -not (Test-PathExistsSmart $p)) {
                                Add-Hit 'Czcionki' 1 $r $n $false $p ('Brak pliku czcionki: ' + $p)
                            }
                        } catch { }
                    }
                } finally { $k.Close() }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'com') {
            Step 'klasy COM (to może chwilę potrwać)'
            foreach ($r in @('HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID', 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\CLSID', 'HKEY_CURRENT_USER\Software\Classes\CLSID')) {
                foreach ($e in [RegDoktor.Reg]::CollectComServers($r)) {
                    if (Test-Cancel) { return }
                    $miss = Get-MissingTarget $e[2]
                    if ($miss) { Add-Hit 'Klasy COM' 2 ($r + '\' + $e[0]) $null $true $miss ('{0}: brak {1}' -f $e[1], $miss) }
                }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'empty') {
            Step 'puste klucze'
            $ex = [string[]]@('Microsoft', 'Classes', 'Policies', 'WOW6432Node', 'AppDataLow', 'Clients', 'RegisteredApplications', 'ODBC', 'Setup', 'DefaultUserEnvironment', 'OEM', 'Partner', 'Google', 'Mozilla')
            foreach ($r in @('HKEY_CURRENT_USER\Software', 'HKEY_LOCAL_MACHINE\SOFTWARE')) {
                foreach ($p in [RegDoktor.Reg]::FindEmptyKeys($r, 2, $ex)) { Add-Hit 'Puste klucze' 2 $p $null $true '' 'Klucz bez wartości i podkluczy' }
            }
        }

        if (Test-Cancel) { return }
        if ($Cats -contains 'services') {
            Step 'usługi i sterowniki'
            $r = 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services'
            foreach ($pair in [RegDoktor.Reg]::CollectChildValues($r, '', 'ImagePath')) {
                $miss = Get-MissingTarget $pair[1]
                if ($miss) { Add-Hit 'Usługi i sterowniki' 3 ($r + '\' + $pair[0]) $null $true $miss ('Usługa {0}: brak {1}' -f $pair[0], $miss) }
            }
        }

        Progress 100 'Skanowanie zakończone'
        if ($found.Count -eq 0) { Ok 'Nie znaleziono osieroconych wpisów. Rejestr jest w dobrym stanie.' }
        else { Ok ('Znaleziono wpisów: {0}' -f $found.Count) }
        foreach ($f in $found) { $f }
    }

    $OrphanCleanCode = {
        $dir = Join-Path $BackupRoot ('Czyszczenie_' + (Get-Stamp))
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $exported = @{}
        $i = 0; $ok = 0; $fail = 0
        $list = @($Items | Sort-Object @{ Expression = { $_.IsWholeKey } }, @{ Expression = { $_.KeyPath.Length }; Descending = $true })
        $total = $list.Count
        foreach ($it in $list) {
            if (Test-Cancel) { break }
            $i++
            Progress ($i * 100.0 / $total) ('Usuwanie {0}/{1}' -f $i, $total)
            try {
                if ($it.IsWholeKey) {
                    if (-not [RegDoktor.Reg]::KeyExists($it.KeyPath)) { $ok++; $it; continue }
                    $file = Join-Path $dir ('{0:D4}_klucz.reg' -f $i)
                    $rc = Export-RegTree $it.KeyPath $file
                    if ($rc -ne 0) { throw 'nie udało się wykonać kopii .reg – wpis pominięto' }
                    [RegDoktor.Reg]::DeleteKeyTree($it.KeyPath)
                } else {
                    if (-not [RegDoktor.Reg]::ValueExists($it.KeyPath, $it.ValueName)) { $ok++; $it; continue }
                    if (-not $exported.ContainsKey($it.KeyPath)) {
                        $file = Join-Path $dir ('{0:D4}_wartosci.reg' -f $i)
                        [void][RegDoktor.Reg]::ExportKeyValues($it.KeyPath, $file)
                        $exported[$it.KeyPath] = $file
                    }
                    [RegDoktor.Reg]::DeleteValue($it.KeyPath, $it.ValueName)
                }
                $ok++
                $it
            } catch {
                $fail++
                $m = $_.Exception.Message
                if ($_.Exception.InnerException) { $m = $_.Exception.InnerException.Message }
                Warn ('Nie usunięto: {0} [{1}] – {2}' -f $it.KeyPath, $it.What, $m)
            }
        }
        if ($ok -gt 0) { Ok ('Usunięto wpisów: {0}. Kopia zapasowa: {1}' -f $ok, $dir) }
        if ($fail -gt 0) { Warn ('Nie udało się usunąć: {0} (brak uprawnień – klucze chronione przez system)' -f $fail) }
        if ($ok -eq 0 -and (Get-ChildItem -LiteralPath $dir -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) { Remove-Item -LiteralPath $dir -Force -ErrorAction SilentlyContinue }
    }

    $HiveAnalyzeCode = {
        $work = Join-Path $env:SystemDrive 'RegDoktor\Analiza'
        if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $work -Force | Out-Null
        $hl = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\hivelist'
        $entries = @()
        foreach ($prop in $hl.PSObject.Properties) {
            $name = $prop.Name; $dev = [string]$prop.Value
            if (-not $dev -or $name -notlike '\REGISTRY\*') { continue }
            if ($name -match '^\\REGISTRY\\MACHINE\\(SOFTWARE|SYSTEM)$') { $entries += ,@('HKLM', $Matches[1], $dev) }
            elseif ($name -match '^\\REGISTRY\\USER\\([^\\]+)$') { $entries += ,@('HKU', $Matches[1], $dev) }
        }
        $i = 0
        foreach ($e in $entries) {
            if (Test-Cancel) { break }
            $i++
            $root = $e[0]; $sub = $e[1]
            $file = [RegDoktor.Native]::DeviceToDos($e[2])
            $h = New-Object RegDoktor.HiveInfo
            $h.Root = $root; $h.SubKey = $sub; $h.FilePath = $file
            $friendly = $sub
            if ($root -eq 'HKU') {
                $isClasses = $sub.EndsWith('_Classes')
                $sidText = $sub -replace '_Classes$', ''
                $acc = $sidText
                switch ($sidText) {
                    '.DEFAULT' { $acc = 'profil domyślny (SYSTEM)' }
                    'S-1-5-19' { $acc = 'LocalService' }
                    'S-1-5-20' { $acc = 'NetworkService' }
                    default { try { $acc = (New-Object System.Security.Principal.SecurityIdentifier $sidText).Translate([System.Security.Principal.NTAccount]).Value } catch { } }
                }
                if ($isClasses) { $friendly = 'UsrClass.dat – ' + $acc } elseif ($sidText -eq '.DEFAULT') { $friendly = 'DEFAULT – ' + $acc } else { $friendly = 'NTUSER.DAT – ' + $acc }
            }
            $h.Name = $friendly
            Progress ($i * 100.0 / ($entries.Count + 1)) ('Analiza: ' + $friendly)
            try { $h.Size = (Get-Item -LiteralPath $file -Force -ErrorAction Stop).Length } catch { $h.Size = 0 }
            $tmp = Join-Path $work ('{0:D2}.hiv' -f $i)
            $rc = [RegDoktor.Native]::SaveHive($root, $sub, $tmp)
            if ($rc -eq 0 -and (Test-Path -LiteralPath $tmp)) {
                $h.Compacted = (Get-Item -LiteralPath $tmp).Length
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                $h.Status = 'Przeanalizowano'
                $h.Selected = ($h.SavingPercent -ge 5 -and ($h.Size - $h.Compacted) -ge 1MB -and $sub -ne 'SYSTEM')
                Log ('{0}: {1} -> {2} (oszczędność {3})' -f $friendly, $h.SizeText, $h.CompactedText, $h.SavingText) 'out'
            } else {
                $h.Status = 'Błąd: ' + (New-Object System.ComponentModel.Win32Exception $rc).Message
                Warn ('{0}: {1}' -f $friendly, $h.Status)
            }
            $h
        }
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        Ok 'Analiza zakończona.'
    }

    $HiveCompactCode = {
        $ts = Get-Stamp
        $ok = 0
        $lines = New-Object System.Collections.Generic.List[string]
        $dirs = @{}
        foreach ($h in $Hives) {
            if (Test-Cancel) { break }
            $drive = [System.IO.Path]::GetPathRoot($h.FilePath)
            if (-not $drive) { Warn ('{0}: nieznana lokalizacja pliku – pominięto' -f $h.Name); continue }
            $dir = Join-Path $drive ('RegDoktor\Kompaktowanie\' + $ts)
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $dirs[$dir] = 1
            $safe = ($h.Root + '_' + $h.SubKey) -replace '[^\w\.-]', '_'
            $new = Join-Path $dir ($safe + '.new')
            $old = Join-Path $dir ($safe + '.old')
            Log ('Kompaktowanie: {0}' -f $h.Name)
            $rc = [RegDoktor.Native]::SaveHive($h.Root, $h.SubKey, $new)
            if ($rc -ne 0) { Err ('{0}: zapis kopii nieudany – {1}' -f $h.Name, (New-Object System.ComponentModel.Win32Exception $rc).Message); continue }
            $newLen = 0
            try { $newLen = (Get-Item -LiteralPath $new -ErrorAction Stop).Length } catch { }
            if ($newLen -le 0 -or ($h.Compacted -gt 0 -and [Math]::Abs($newLen - $h.Compacted) -gt 0.25 * $h.Compacted)) {
                Err ('{0}: rozmiar nowej kopii ({1}) odbiega od analizy ({2}) – pominięto ze względów bezpieczeństwa. Wykonaj analizę ponownie.' -f $h.Name, (Format-Size $newLen), $h.CompactedText)
                Remove-Item -LiteralPath $new -Force -ErrorAction SilentlyContinue
                continue
            }
            $rc = [RegDoktor.Native]::ReplaceHive($h.Root, $h.SubKey, $new, $old)
            if ($rc -ne 0) {
                Err ('{0}: podmiana nieudana – {1}' -f $h.Name, (New-Object System.ComponentModel.Win32Exception $rc).Message)
                Remove-Item -LiteralPath $new -Force -ErrorAction SilentlyContinue
                continue
            }
            $ok++
            Ok ('{0}: zaplanowano podmianę ({1} -> {2})' -f $h.Name, $h.SizeText, $h.CompactedText)
            $lines.Add(('copy /y "{0}" "{1}"' -f $old, $h.FilePath))
        }
        foreach ($d in $dirs.Keys) {
            $txt = @"
RegDoktor 11 - kompaktowanie rejestru: $(Get-Date)

Oryginalne pliki gałęzi zostały zachowane jako *.old w tym folderze.
Jeśli po restarcie Windows nie uruchomi się poprawnie:
 1. Uruchom Środowisko odzyskiwania (przytrzymaj Shift i kliknij Uruchom ponownie,
    albo 3x przerwij uruchamianie) > Rozwiązywanie problemów > Opcje zaawansowane > Wiersz polecenia.
 2. W WinRE litera dysku z Windows może być inna (sprawdź: dir C:\ , dir D:\ ) - popraw ją w poleceniach.
 3. Wykonaj:

$($lines -join "`r`n")

 4. Wpisz exit i uruchom komputer ponownie.
Możesz też przywrócić punkt przywracania systemu utworzony przed operacją.
Gdy system działa poprawnie, ten folder można usunąć.
"@
            Set-Content -LiteralPath (Join-Path $d 'PRZYWRACANIE.txt') -Value $txt -Encoding UTF8
            Log ('Instrukcja przywrócenia: ' + (Join-Path $d 'PRZYWRACANIE.txt')) 'out'
        }
        if ($ok -gt 0) { Ok ('Zaplanowano kompaktowanie {0} gałęzi – uruchom komputer ponownie.' -f $ok) }
        @{ Scheduled = $ok }
    }

# ============================================================== XAML
$MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RegDoktor 11 — rejestr, naprawy i optymalizacja Windows"
        Width="1380" Height="900" MinWidth="1040" MinHeight="680"
        WindowStartupLocation="CenterScreen"
        Background="#15161A" Foreground="#E6E7EB"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
  <Window.Resources>
    <SolidColorBrush x:Key="TextBrush" Color="#E6E7EB"/>
    <SolidColorBrush x:Key="MutedBrush" Color="#9AA0AD"/>
    <SolidColorBrush x:Key="PanelBrush" Color="#1D1F25"/>
    <SolidColorBrush x:Key="Panel2Brush" Color="#23262E"/>
    <SolidColorBrush x:Key="LineBrush" Color="#2F333D"/>
    <SolidColorBrush x:Key="InputBrush" Color="#1A1C22"/>
    <SolidColorBrush x:Key="AccentBrush" Color="#4F8CFF"/>

    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#2E4A7A"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightBrushKey}" Color="#2A3446"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightTextBrushKey}" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#2A3446"/>

    <!-- thin dark scrollbars -->
    <Style x:Key="ThumbStyle" TargetType="Thumb">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border x:Name="th" Background="#3A3F4B" CornerRadius="4" Margin="2"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="th" Property="Background" Value="#545B6B"/></Trigger>
              <Trigger Property="IsDragging" Value="True"><Setter TargetName="th" Property="Background" Value="#6B7385"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Width" Value="11"/>
      <Setter Property="MinWidth" Value="11"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid Background="{TemplateBinding Background}">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.Thumb><Thumb Style="{StaticResource ThumbStyle}"/></Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="11"/>
          <Setter Property="MinHeight" Value="11"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Grid Background="{TemplateBinding Background}">
                  <Track x:Name="PART_Track" IsDirectionReversed="False">
                    <Track.Thumb><Thumb Style="{StaticResource ThumbStyle}"/></Track.Thumb>
                  </Track>
                </Grid>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- buttons -->
    <Style TargetType="Button">
      <Setter Property="Foreground" Value="#E6E7EB"/>
      <Setter Property="Background" Value="#2A2D36"/>
      <Setter Property="BorderBrush" Value="#3A3F4B"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Grid>
              <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6"/>
              <Border x:Name="ov" Background="#18FFFFFF" CornerRadius="6" Visibility="Collapsed"/>
              <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ov" Property="Visibility" Value="Visible"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="ov" Property="Background" Value="#22000000"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="PrimaryBtn" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#3B82F6"/>
      <Setter Property="BorderBrush" Value="#3B82F6"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="DangerBtn" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#C2410C"/>
      <Setter Property="BorderBrush" Value="#C2410C"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="ChipBtn" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Padding" Value="12,5"/>
      <Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="FontSize" Value="12.5"/>
    </Style>

    <!-- colourful repair tiles (Background = tint, BorderBrush = category colour) -->
    <Style x:Key="ActionBtn" TargetType="Button">
      <Setter Property="Foreground" Value="#F1F2F5"/>
      <Setter Property="Width" Value="266"/>
      <Setter Property="MinHeight" Value="82"/>
      <Setter Property="Margin" Value="0,0,10,10"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="VerticalAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Grid>
              <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="#1EFFFFFF" BorderThickness="1" CornerRadius="9"/>
              <Border Width="4" HorizontalAlignment="Left" Margin="1,12,0,12" CornerRadius="2" Background="{TemplateBinding BorderBrush}"/>
              <Border x:Name="ov" CornerRadius="9" Background="#14FFFFFF" Visibility="Collapsed"/>
              <ContentPresenter Margin="16,10,12,10" HorizontalAlignment="Stretch" VerticalAlignment="Top"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ov" Property="Visibility" Value="Visible"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="{Binding BorderBrush, RelativeSource={RelativeSource TemplatedParent}}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="ov" Property="Background" Value="#26000000"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- tabs -->
    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabControl">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <Border Background="#1A1B20" BorderBrush="#2A2D36" BorderThickness="0,0,0,1">
                <TabPanel IsItemsHost="True" Margin="14,8,14,8" Background="Transparent"/>
              </Border>
              <ContentPresenter Grid.Row="1" ContentSource="SelectedContent"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="#AEB3BF"/>
      <Setter Property="FontSize" Value="13.5"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" CornerRadius="7" Padding="14,7" Margin="0,0,6,0">
              <ContentPresenter ContentSource="Header" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#24272F"/></Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#26395C"/>
                <Setter Property="Foreground" Value="#FFFFFF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- inputs -->
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#1A1C22"/>
      <Setter Property="Foreground" Value="#E6E7EB"/>
      <Setter Property="BorderBrush" Value="#3A3F4B"/>
      <Setter Property="CaretBrush" Value="#E6E7EB"/>
      <Setter Property="SelectionBrush" Value="#4F8CFF"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#D7DAE1"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,12,0"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="#D7DAE1"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,12,0"/>
    </Style>
    <Style TargetType="ProgressBar">
      <Setter Property="Foreground" Value="#4F8CFF"/>
      <Setter Property="Background" Value="#2A2D36"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Height" Value="6"/>
    </Style>
    <Style TargetType="GridSplitter">
      <Setter Property="Background" Value="#2A2D36"/>
    </Style>

    <!-- data grids -->
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="#1A1C22"/>
      <Setter Property="Foreground" Value="#E6E7EB"/>
      <Setter Property="BorderBrush" Value="#2F333D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="#1A1C22"/>
      <Setter Property="AlternatingRowBackground" Value="#1E2027"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#262931"/>
      <Setter Property="VerticalGridLinesBrush" Value="#262931"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="CanUserDeleteRows" Value="False"/>
      <Setter Property="CanUserResizeRows" Value="False"/>
      <Setter Property="SelectionUnit" Value="FullRow"/>
      <Setter Property="RowHeaderWidth" Value="0"/>
      <Setter Property="EnableRowVirtualization" Value="True"/>
      <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True"/>
      <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling"/>
      <Setter Property="FontSize" Value="12.5"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#23262E"/>
      <Setter Property="Foreground" Value="#AEB3BF"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="#2F333D"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="7,3"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#2E4A7A"/>
          <Setter Property="Foreground" Value="#FFFFFF"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="#1D1F25"/>
      <Setter Property="BorderBrush" Value="#2A2D36"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="14,10"/>
    </Style>
    <Style x:Key="InfoCard" TargetType="Border" BasedOn="{StaticResource Card}">
      <Setter Property="Margin" Value="0,0,10,0"/>
      <Setter Property="MinWidth" Value="170"/>
      <Setter Property="Padding" Value="14,8"/>
    </Style>
    <Style x:Key="InfoTitle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#8C93A1"/>
      <Setter Property="FontSize" Value="11.5"/>
    </Style>
    <Style x:Key="InfoValue" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#F1F2F5"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,2,0,0"/>
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#9AA0AD"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="210" MinHeight="90"/>
    </Grid.RowDefinitions>

    <!-- ===================== HEADER ===================== -->
    <Border Grid.Row="0" Background="#1A1B20" Padding="16,10">
      <DockPanel LastChildFill="False">
        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
          <Border Width="40" Height="40" CornerRadius="10">
            <Border.Background>
              <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                <GradientStop Color="#3B82F6" Offset="0"/>
                <GradientStop Color="#8B5CF6" Offset="0.55"/>
                <GradientStop Color="#EC4899" Offset="1"/>
              </LinearGradientBrush>
            </Border.Background>
            <TextBlock Text="R" FontSize="22" FontWeight="Bold" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
            <TextBlock Text="RegDoktor 11" FontSize="19" FontWeight="SemiBold" Foreground="#FFFFFF"/>
            <TextBlock Text="Przeglądarka rejestru  •  Osierocone wpisy  •  Defragmentacja  •  Naprawy systemu, sieci i pamięci" FontSize="12" Foreground="#9AA0AD"/>
          </StackPanel>
        </StackPanel>
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <Border x:Name="bdAdmin" CornerRadius="10" Padding="10,3" Background="#163A2B" Margin="0,0,14,0" VerticalAlignment="Center">
            <TextBlock x:Name="txtAdmin" Text="Administrator" Foreground="#34D399" FontSize="12" FontWeight="SemiBold"/>
          </Border>
          <CheckBox x:Name="chkRestorePoint" IsChecked="True" Content="Punkt przywracania przed ryzykownymi zmianami" Margin="0,0,14,0"
                    ToolTip="Przed akcjami oznaczonymi jako ryzykowne zostanie automatycznie utworzony punkt przywracania systemu."/>
          <Button x:Name="btnRestorePoint" Content="Utwórz punkt przywracania"/>
          <Button x:Name="btnBackups" Content="Folder kopii zapasowych" Margin="0"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- ===================== TABS ===================== -->
    <TabControl x:Name="tabMain" Grid.Row="1">

      <!-- ===== TAB: repairs ===== -->
      <TabItem x:Name="tabFix" Header="⚙  Naprawy i optymalizacja">
        <Grid TextElement.Foreground="#E6E7EB" TextElement.FontSize="13">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="16,12,16,0">
            <Border Style="{StaticResource InfoCard}"><StackPanel><TextBlock Style="{StaticResource InfoTitle}" Text="System"/><TextBlock x:Name="txtInfoOS" Style="{StaticResource InfoValue}" Text="…"/></StackPanel></Border>
            <Border Style="{StaticResource InfoCard}"><StackPanel><TextBlock Style="{StaticResource InfoTitle}" Text="Czas pracy"/><TextBlock x:Name="txtInfoUptime" Style="{StaticResource InfoValue}" Text="…"/></StackPanel></Border>
            <Border Style="{StaticResource InfoCard}"><StackPanel><TextBlock Style="{StaticResource InfoTitle}" Text="Pamięć RAM"/><TextBlock x:Name="txtInfoRam" Style="{StaticResource InfoValue}" Text="…"/></StackPanel></Border>
            <Border Style="{StaticResource InfoCard}"><StackPanel><TextBlock Style="{StaticResource InfoTitle}" Text="Dysk systemowy"/><TextBlock x:Name="txtInfoDisk" Style="{StaticResource InfoValue}" Text="…"/></StackPanel></Border>
            <Border Style="{StaticResource InfoCard}"><StackPanel><TextBlock Style="{StaticResource InfoTitle}" Text="Ostatni rozruch"/><TextBlock x:Name="txtInfoBoot" Style="{StaticResource InfoValue}" Text="…"/></StackPanel></Border>
            <Button x:Name="btnInfoRefresh" Content="↻ Odśwież" VerticalAlignment="Center"/>
          </StackPanel>
          <Grid Grid.Row="1" Margin="16,12,16,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="300"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid VerticalAlignment="Top">
              <TextBox x:Name="txtFilter" Height="32"/>
              <TextBlock x:Name="lblFilterHint" Text="Szukaj naprawy (np. dns, dysk, menu start, wifi)…" Foreground="#6E7483" Margin="10,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
            </Grid>
            <WrapPanel x:Name="pnlCatChips" Grid.Column="1" Margin="14,0,0,0"/>
          </Grid>
          <ScrollViewer x:Name="svActions" Grid.Row="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="16,4,6,12">
            <StackPanel x:Name="pnlActions"/>
          </ScrollViewer>
        </Grid>
      </TabItem>

      <!-- ===== TAB: registry browser ===== -->
      <TabItem x:Name="tabReg" Header="☰  Przeglądarka rejestru">
        <Grid TextElement.Foreground="#E6E7EB" TextElement.FontSize="13" Margin="14,10,14,10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="5"/>
            <RowDefinition Height="230" MinHeight="120"/>
          </Grid.RowDefinitions>
          <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Adres:" Foreground="#9AA0AD" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox x:Name="txtAddress" Grid.Column="1" Height="32" FontFamily="Cascadia Mono, Consolas" FontSize="12.5"
                     ToolTip="Wklej ścieżkę (np. HKLM\SOFTWARE\Microsoft lub Komputer\HKEY_CURRENT_USER\...) i naciśnij Enter"/>
            <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="8,0,0,0">
              <Button x:Name="btnGo" Content="Przejdź" Style="{StaticResource PrimaryBtn}"/>
              <Button x:Name="btnUp" Content="↑ W górę"/>
              <Button x:Name="btnRegRefresh" Content="↻" ToolTip="Odśwież (F5)" Margin="0"/>
            </StackPanel>
          </Grid>
          <WrapPanel Grid.Row="1" Margin="0,8,0,8">
            <Button x:Name="btnNewKey" Content="+ Nowy klucz"/>
            <Button x:Name="btnNewValue" Content="+ Nowa wartość"/>
            <Button x:Name="btnEditValue" Content="✎ Edytuj wartość"/>
            <Button x:Name="btnDeleteReg" Content="✖ Usuń" ToolTip="Usuwa zaznaczone wartości, a jeśli żadna nie jest zaznaczona – zaznaczony klucz. Zawsze z kopią .reg."/>
            <Button x:Name="btnExportReg" Content="↓ Eksportuj .reg"/>
            <Button x:Name="btnImportReg" Content="↑ Importuj .reg"/>
            <Button x:Name="btnCopyPath" Content="Kopiuj ścieżkę"/>
            <Button x:Name="btnOpenRegedit" Content="Otwórz w Regedit"/>
          </WrapPanel>
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="400" MinWidth="200"/>
              <ColumnDefinition Width="5"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TreeView x:Name="tvReg" Background="#1A1C22" BorderBrush="#2F333D" Foreground="#E6E7EB"
                      VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling">
              <TreeView.ItemContainerStyle>
                <Style TargetType="TreeViewItem">
                  <Setter Property="IsExpanded" Value="{Binding IsExpanded, Mode=TwoWay}"/>
                  <Setter Property="IsSelected" Value="{Binding IsSelected, Mode=TwoWay}"/>
                  <Setter Property="Foreground" Value="#E6E7EB"/>
                  <Setter Property="Padding" Value="2,1"/>
                </Style>
              </TreeView.ItemContainerStyle>
              <TreeView.ItemTemplate>
                <HierarchicalDataTemplate ItemsSource="{Binding Children}">
                  <StackPanel Orientation="Horizontal">
                    <Path Data="M0,1.5 L4.5,1.5 L6,3 L13,3 L13,11 L0,11 Z" Fill="#E8B04A" Width="13" Height="11" Stretch="Fill" Margin="0,1,6,0" VerticalAlignment="Center"/>
                    <TextBlock Text="{Binding Name}" VerticalAlignment="Center"/>
                  </StackPanel>
                </HierarchicalDataTemplate>
              </TreeView.ItemTemplate>
            </TreeView>
            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" ResizeDirection="Columns" Background="Transparent"/>
            <DataGrid x:Name="dgValues" Grid.Column="2" IsReadOnly="True" SelectionMode="Extended">
              <DataGrid.Columns>
                <DataGridTemplateColumn Header="" Width="34">
                  <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                      <TextBlock x:Name="ic" Text="{Binding Icon}" FontWeight="Bold" FontSize="11" Foreground="#6BA6E5" HorizontalAlignment="Center"/>
                      <DataTemplate.Triggers>
                        <DataTrigger Binding="{Binding IsString}" Value="True"><Setter TargetName="ic" Property="Foreground" Value="#E5826B"/></DataTrigger>
                      </DataTemplate.Triggers>
                    </DataTemplate>
                  </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Nazwa" Binding="{Binding DisplayName, Mode=OneWay}" Width="240"/>
                <DataGridTextColumn Header="Typ" Binding="{Binding KindName, Mode=OneWay}" Width="125"/>
                <DataGridTextColumn Header="Dane" Binding="{Binding Data, Mode=OneWay}" Width="*"/>
              </DataGrid.Columns>
            </DataGrid>
          </Grid>
          <GridSplitter Grid.Row="3" Height="5" HorizontalAlignment="Stretch" ResizeDirection="Rows" Background="Transparent"/>
          <Border Grid.Row="4" Style="{StaticResource Card}" Padding="10">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <DockPanel LastChildFill="True" Margin="0,0,0,8">
                <TextBlock DockPanel.Dock="Left" Text="Szukaj:" VerticalAlignment="Center" Foreground="#9AA0AD" Margin="0,0,8,0"/>
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" Margin="10,0,0,0">
                  <CheckBox x:Name="chkSKeys" Content="Klucze" IsChecked="True"/>
                  <CheckBox x:Name="chkSNames" Content="Nazwy wartości" IsChecked="True"/>
                  <CheckBox x:Name="chkSData" Content="Dane" IsChecked="True"/>
                  <RadioButton x:Name="rbScopeSel" Content="W zaznaczonym kluczu" IsChecked="True" GroupName="scope"/>
                  <RadioButton x:Name="rbScopeAll" Content="HKCU + HKLM" GroupName="scope"/>
                  <Button x:Name="btnSearch" Content="Szukaj" Style="{StaticResource PrimaryBtn}"/>
                  <Button x:Name="btnSearchStop" Content="Stop" IsEnabled="False" Margin="0"/>
                </StackPanel>
                <TextBox x:Name="txtSearch" Height="30"/>
              </DockPanel>
              <Grid Grid.Row="1">
                <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                <DataGrid x:Name="dgSearch" IsReadOnly="True" SelectionMode="Single">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Ścieżka klucza" Binding="{Binding Path, Mode=OneWay}" Width="3*"/>
                    <DataGridTextColumn Header="Trafienie" Binding="{Binding Where, Mode=OneWay}" Width="110"/>
                    <DataGridTextColumn Header="Wartość" Binding="{Binding ValueName, Mode=OneWay}" Width="*"/>
                    <DataGridTextColumn Header="Dane" Binding="{Binding Data, Mode=OneWay}" Width="2*"/>
                  </DataGrid.Columns>
                </DataGrid>
                <TextBlock x:Name="txtSearchStatus" Grid.Row="1" Style="{StaticResource Hint}" Margin="2,6,0,0" Text="Dwuklik na wyniku przenosi do klucza."/>
              </Grid>
            </Grid>
          </Border>
          <TextBlock x:Name="txtRegStatus" Grid.Row="2" Visibility="Collapsed"/>
        </Grid>
      </TabItem>

      <!-- ===== TAB: orphans ===== -->
      <TabItem x:Name="tabClean" Header="♻  Osierocone wpisy">
        <Grid TextElement.Foreground="#E6E7EB" TextElement.FontSize="13" Margin="14,10,14,10">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="330"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Border Style="{StaticResource Card}" Margin="0,0,12,0">
            <DockPanel>
              <TextBlock DockPanel.Dock="Top" Text="Co skanować" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,4"/>
              <TextBlock DockPanel.Dock="Top" Style="{StaticResource Hint}" Margin="0,0,0,10"
                         Text="Szukane są wyłącznie wpisy wskazujące na nieistniejące pliki/foldery (na dyskach lokalnych) lub zupełnie puste klucze. Przed usunięciem każdy wpis trafia do kopii .reg."/>
              <StackPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
                <Button x:Name="btnScan" Content="Skanuj rejestr" Style="{StaticResource PrimaryBtn}" Height="38" Margin="0,0,0,8" HorizontalAlignment="Stretch"/>
                <TextBlock Style="{StaticResource Hint}" FontSize="11.5">
                  <Run Foreground="#34D399" Text="■ Niskie"/> – bezpieczne do usunięcia
                  <LineBreak/><Run Foreground="#FBBF24" Text="■ Średnie"/> – zwykle bezpieczne, sprawdź listę
                  <LineBreak/><Run Foreground="#F87171" Text="■ Wysokie"/> – tylko jeśli wiesz, co robisz
                </TextBlock>
              </StackPanel>
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="pnlScanCats"/>
              </ScrollViewer>
            </DockPanel>
          </Border>
          <Grid Grid.Column="1">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <DockPanel Margin="0,0,0,8" LastChildFill="True">
              <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <Button x:Name="btnRestoreBackup" Content="↶ Przywróć z kopii…"/>
                <Button x:Name="btnClean" Content="Usuń zaznaczone (z kopią)" Style="{StaticResource DangerBtn}" Margin="0" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="btnSelSafe" Content="Zaznacz bezpieczne"/>
                <Button x:Name="btnSelAll" Content="Zaznacz wszystko"/>
                <Button x:Name="btnSelNone" Content="Odznacz"/>
                <TextBlock x:Name="txtScanSummary" Style="{StaticResource Hint}" VerticalAlignment="Center" Margin="8,0,0,0" Text="Uruchom skanowanie, aby zobaczyć wyniki."/>
              </StackPanel>
            </DockPanel>
            <DataGrid x:Name="dgClean" Grid.Row="1" SelectionMode="Extended" IsReadOnly="False">
              <DataGrid.Columns>
                <DataGridTemplateColumn Header="✓" Width="36">
                  <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                      <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" Margin="0"/>
                    </DataTemplate>
                  </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Ryzyko" Width="78" SortMemberPath="RiskLevel">
                  <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                      <TextBlock x:Name="rk" Text="{Binding Risk}" Foreground="#34D399" FontWeight="SemiBold"/>
                      <DataTemplate.Triggers>
                        <DataTrigger Binding="{Binding RiskLevel}" Value="2"><Setter TargetName="rk" Property="Foreground" Value="#FBBF24"/></DataTrigger>
                        <DataTrigger Binding="{Binding RiskLevel}" Value="3"><Setter TargetName="rk" Property="Foreground" Value="#F87171"/></DataTrigger>
                      </DataTemplate.Triggers>
                    </DataTemplate>
                  </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Kategoria" Binding="{Binding Category, Mode=OneWay}" Width="150" IsReadOnly="True"/>
                <DataGridTextColumn Header="Klucz" Binding="{Binding KeyPath, Mode=OneWay}" Width="3*" IsReadOnly="True"/>
                <DataGridTextColumn Header="Wartość" Binding="{Binding What, Mode=OneWay}" Width="*" IsReadOnly="True"/>
                <DataGridTextColumn Header="Brakujący element / szczegóły" Binding="{Binding Detail, Mode=OneWay}" Width="2*" IsReadOnly="True"/>
              </DataGrid.Columns>
            </DataGrid>
          </Grid>
        </Grid>
      </TabItem>

      <!-- ===== TAB: hive compaction ===== -->
      <TabItem x:Name="tabHive" Header="▦  Defragmentacja rejestru">
        <Grid TextElement.Foreground="#E6E7EB" TextElement.FontSize="13" Margin="14,10,14,10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Border Style="{StaticResource Card}" Margin="0,0,0,10">
            <StackPanel>
              <TextBlock Text="Kompaktowanie (defragmentacja) gałęzi rejestru" FontSize="15" FontWeight="SemiBold"/>
              <TextBlock Style="{StaticResource Hint}" Margin="0,6,0,0">
                Pliki gałęzi (SOFTWARE, SYSTEM, NTUSER.DAT…) z czasem zawierają puste miejsca po usuniętych kluczach. Analiza zapisuje skompaktowaną kopię każdej gałęzi
                (nic nie zmieniając) i pokazuje realny zysk. „Kompaktuj” podmienia pliki gałęzi przy najbliższym restarcie (to samo API, którego używa Windows – RegReplaceKey).
                <LineBreak/>Windows (od wersji 8) sam kompaktuje gałęzie przy uruchamianiu, gdy zysk jest duży – dlatego oszczędność zwykle wynosi od kilku do kilkudziesięciu MB. Gałąź SYSTEM nie jest zaznaczana domyślnie. Oryginalne pliki zostają zachowane (.old) wraz z instrukcją przywrócenia.
              </TextBlock>
            </StackPanel>
          </Border>
          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <Button x:Name="btnHiveAnalyze" Content="Analizuj" Style="{StaticResource PrimaryBtn}"/>
            <Button x:Name="btnHiveCompact" Content="Kompaktuj zaznaczone (restart)" Style="{StaticResource DangerBtn}" IsEnabled="False"/>
            <TextBlock x:Name="txtHiveSummary" Style="{StaticResource Hint}" VerticalAlignment="Center" Margin="8,0,0,0" Text="Najpierw wykonaj analizę."/>
          </StackPanel>
          <DataGrid x:Name="dgHives" Grid.Row="2" IsReadOnly="False" SelectionMode="Single">
            <DataGrid.Columns>
              <DataGridTemplateColumn Header="✓" Width="36">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" Margin="0"/>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Header="Gałąź" Binding="{Binding Name, Mode=OneWay}" Width="230" IsReadOnly="True"/>
              <DataGridTextColumn Header="Plik" Binding="{Binding FilePath, Mode=OneWay}" Width="2*" IsReadOnly="True"/>
              <DataGridTextColumn Header="Rozmiar" Binding="{Binding SizeText, Mode=OneWay}" Width="100" IsReadOnly="True"/>
              <DataGridTextColumn Header="Po kompaktowaniu" Binding="{Binding CompactedText, Mode=OneWay}" Width="135" IsReadOnly="True"/>
              <DataGridTextColumn Header="Oszczędność" Binding="{Binding SavingText, Mode=OneWay}" Width="150" IsReadOnly="True"/>
              <DataGridTextColumn Header="Status" Binding="{Binding Status, Mode=OneWay}" Width="*" IsReadOnly="True"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </TabItem>

      <!-- ===== TAB: startup ===== -->
      <TabItem x:Name="tabStart" Header="⚡  Autostart">
        <Grid TextElement.Foreground="#E6E7EB" TextElement.FontSize="13" Margin="14,10,14,10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <DockPanel Margin="0,0,0,8">
            <StackPanel Orientation="Horizontal">
              <Button x:Name="btnStartRefresh" Content="↻ Odśwież"/>
              <Button x:Name="btnStartEnable" Content="✔ Włącz"/>
              <Button x:Name="btnStartDisable" Content="✖ Wyłącz"/>
              <Button x:Name="btnStartDelete" Content="Usuń wpis (z kopią)"/>
              <Button x:Name="btnStartOpen" Content="Pokaż plik"/>
              <TextBlock x:Name="txtStartSummary" Style="{StaticResource Hint}" VerticalAlignment="Center" Margin="8,0,0,0"
                         Text="Wyłączanie działa tak samo jak w Menedżerze zadań – w każdej chwili można je cofnąć."/>
            </StackPanel>
          </DockPanel>
          <DataGrid x:Name="dgStartup" Grid.Row="1" IsReadOnly="True" SelectionMode="Extended">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Nazwa" Binding="{Binding Name, Mode=OneWay}" Width="220"/>
              <DataGridTemplateColumn Header="Stan" Width="170" SortMemberPath="Status">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <TextBlock x:Name="st" Text="{Binding Status}" Foreground="#34D399" FontWeight="SemiBold"/>
                    <DataTemplate.Triggers>
                      <DataTrigger Binding="{Binding Enabled}" Value="False"><Setter TargetName="st" Property="Foreground" Value="#9AA0AD"/></DataTrigger>
                      <DataTrigger Binding="{Binding Missing}" Value="True"><Setter TargetName="st" Property="Foreground" Value="#F87171"/></DataTrigger>
                    </DataTemplate.Triggers>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Header="Zakres" Binding="{Binding Scope, Mode=OneWay}" Width="150"/>
              <DataGridTextColumn Header="Polecenie" Binding="{Binding Command, Mode=OneWay}" Width="3*"/>
              <DataGridTextColumn Header="Lokalizacja" Binding="{Binding Location, Mode=OneWay}" Width="2*"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </TabItem>
    </TabControl>

    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch" ResizeDirection="Rows"/>

    <!-- ===================== LOG ===================== -->
    <Border Grid.Row="3" Background="#111216">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid Margin="14,7,14,5">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="220"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="Dziennik" FontWeight="SemiBold" Foreground="#C9CDD6" VerticalAlignment="Center" Margin="0,0,14,0"/>
          <TextBlock x:Name="txtStatus" Grid.Column="1" Text="Gotowy." Foreground="#9AA0AD" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
          <ProgressBar x:Name="pbProgress" Grid.Column="2" Minimum="0" Maximum="100" Value="0" Margin="10,0,10,0" VerticalAlignment="Center"/>
          <StackPanel Grid.Column="3" Orientation="Horizontal">
            <Button x:Name="btnCancel" Content="✖ Anuluj zadanie" IsEnabled="False" Padding="10,3"/>
            <Button x:Name="btnLogCopy" Content="Kopiuj" Padding="10,3"/>
            <Button x:Name="btnLogSave" Content="Zapisz…" Padding="10,3"/>
            <Button x:Name="btnLogClear" Content="Wyczyść" Padding="10,3" Margin="0"/>
          </StackPanel>
        </Grid>
        <ListBox x:Name="lstLog" Grid.Row="1" Background="#0E0F12" BorderThickness="0" Margin="8,0,8,8"
                 VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling"
                 ScrollViewer.HorizontalScrollBarVisibility="Auto" SelectionMode="Extended">
          <ListBox.ItemTemplate>
            <DataTemplate>
              <TextBlock x:Name="t" Text="{Binding Display}" FontFamily="Cascadia Mono, Consolas" FontSize="12" Foreground="#C9CDD6"/>
              <DataTemplate.Triggers>
                <DataTrigger Binding="{Binding Kind}" Value="ok"><Setter TargetName="t" Property="Foreground" Value="#34D399"/></DataTrigger>
                <DataTrigger Binding="{Binding Kind}" Value="warn"><Setter TargetName="t" Property="Foreground" Value="#FBBF24"/></DataTrigger>
                <DataTrigger Binding="{Binding Kind}" Value="err"><Setter TargetName="t" Property="Foreground" Value="#F87171"/></DataTrigger>
                <DataTrigger Binding="{Binding Kind}" Value="cmd"><Setter TargetName="t" Property="Foreground" Value="#67E8F9"/></DataTrigger>
                <DataTrigger Binding="{Binding Kind}" Value="out"><Setter TargetName="t" Property="Foreground" Value="#A3A9B6"/></DataTrigger>
                <DataTrigger Binding="{Binding Kind}" Value="head">
                  <Setter TargetName="t" Property="Foreground" Value="#93C5FD"/>
                  <Setter TargetName="t" Property="FontWeight" Value="Bold"/>
                </DataTrigger>
              </DataTemplate.Triggers>
            </DataTemplate>
          </ListBox.ItemTemplate>
        </ListBox>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$ValueDialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Wartość rejestru" Width="620" Height="470" MinWidth="440" MinHeight="330"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="#1D1F25" Foreground="#E6E7EB" FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13">
  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Text="Nazwa wartości" Foreground="#9AA0AD"/>
    <TextBox x:Name="txtName" Grid.Row="1" Height="30" Margin="0,4,0,12"/>
    <TextBlock Grid.Row="2" Text="Typ" Foreground="#9AA0AD"/>
    <WrapPanel x:Name="pnlTypes" Grid.Row="3" Margin="0,6,0,12">
      <RadioButton x:Name="rbSZ" Content="REG_SZ" GroupName="t" Margin="0,0,14,4"/>
      <RadioButton x:Name="rbEXP" Content="REG_EXPAND_SZ" GroupName="t" Margin="0,0,14,4"/>
      <RadioButton x:Name="rbDW" Content="REG_DWORD" GroupName="t" Margin="0,0,14,4"/>
      <RadioButton x:Name="rbQW" Content="REG_QWORD" GroupName="t" Margin="0,0,14,4"/>
      <RadioButton x:Name="rbMULTI" Content="REG_MULTI_SZ" GroupName="t" Margin="0,0,14,4"/>
      <RadioButton x:Name="rbBIN" Content="REG_BINARY" GroupName="t" Margin="0,0,14,4"/>
    </WrapPanel>
    <TextBlock x:Name="txtHint" Grid.Row="4" Foreground="#9AA0AD" TextWrapping="Wrap" FontSize="12"/>
    <TextBox x:Name="txtData" Grid.Row="5" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
             VerticalContentAlignment="Top" FontFamily="Cascadia Mono, Consolas" Margin="0,6,0,14"/>
    <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="btnOk" Content="Zapisz" IsDefault="True" MinWidth="100"/>
      <Button x:Name="btnCancelDlg" Content="Anuluj" IsCancel="True" MinWidth="100" Margin="0"/>
    </StackPanel>
  </Grid>
</Window>
'@

$InputDialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RegDoktor" Width="520" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="#1D1F25" Foreground="#E6E7EB" FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13">
  <StackPanel Margin="18">
    <TextBlock x:Name="txtPrompt" TextWrapping="Wrap" Foreground="#C9CDD6"/>
    <TextBox x:Name="txtInput" Height="30" Margin="0,10,0,16"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="btnOk" Content="OK" IsDefault="True" MinWidth="100"/>
      <Button x:Name="btnCancelDlg" Content="Anuluj" IsCancel="True" MinWidth="100" Margin="0"/>
    </StackPanel>
  </StackPanel>
</Window>
'@

# ============================================================== UI helpers
function Import-Xaml([string]$Text) {
    $xml = [xml]$Text
    $win = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xml))
    $h = @{ Window = $win }
    foreach ($m in [regex]::Matches($Text, 'x:Name="([^"]+)"')) {
        $n = $m.Groups[1].Value
        $o = $win.FindName($n)
        if ($null -ne $o) { $h[$n] = $o }
    }
    return $h
}

function Get-ExceptionChain($Exception) {
    $parts = New-Object System.Collections.Generic.List[string]
    $e = $Exception
    while ($e) {
        $line = '{0}: {1}' -f $e.GetType().Name, $e.Message
        if ($e -is [System.Windows.Markup.XamlParseException] -and $e.LineNumber -gt 0) { $line += (' [XAML linia {0}, poz. {1}]' -f $e.LineNumber, $e.LinePosition) }
        $parts.Add($line)
        $e = $e.InnerException
    }
    return ($parts -join "`n -> ")
}

# XAML elements that differ most between Windows/.NET builds. If the window fails to load,
# they are removed one by one (cumulatively) until it loads - the app then works with a plainer look.
$XamlFallbacks = @(
    @{ Name = 'kolory zaznaczenia (SystemColors)'; XPath = "//*[local-name()='SolidColorBrush'][starts-with(@*[local-name()='Key'], '{x:Static')]" }
    @{ Name = 'podświetlenie kafelków (Binding w wyzwalaczu)'; XPath = "//*[local-name()='Setter'][starts-with(@Value, '{Binding BorderBrush')]" }
    @{ Name = 'styl pasków przewijania'; XPath = "//*[local-name()='Style'][@TargetType='ScrollBar' or @TargetType='Thumb']" }
    @{ Name = 'TextOptions / UseLayoutRounding'; Attr = @('TextOptions.TextFormattingMode', 'UseLayoutRounding') }
    @{ Name = 'TextElement.* na kontenerach'; Attr = @('TextElement.Foreground', 'TextElement.FontSize') }
    @{ Name = 'wirtualizacja list'; Attr = @('VirtualizingPanel.IsVirtualizing', 'VirtualizingPanel.VirtualizationMode'); XPath = "//*[local-name()='Setter'][starts-with(@Property, 'VirtualizingPanel.')]" }
    @{ Name = 'styl komórek i nagłówków tabel'; XPath = "//*[local-name()='Style'][@TargetType='DataGridCell' or @TargetType='DataGridColumnHeader']" }
    @{ Name = 'styl zakładek'; XPath = "//*[local-name()='Style'][@TargetType='TabControl' or @TargetType='TabItem']" }
    @{ Name = 'styl drzewa rejestru'; XPath = "//*[local-name()='Setter'][@Property='IsExpanded' or @Property='IsSelected']" }
)

function Import-MainXaml([string]$Text) {
    $first = $null
    try { return (Import-Xaml $Text) } catch { $first = $_ }
    $State.XamlError = Get-ExceptionChain $first.Exception
    $doc = [xml]$Text
    $applied = New-Object System.Collections.Generic.List[string]
    foreach ($fb in $XamlFallbacks) {
        if ($fb.XPath) { foreach ($n in @($doc.SelectNodes($fb.XPath))) { [void]$n.ParentNode.RemoveChild($n) } }
        if ($fb.Attr) {
            foreach ($el in @($doc.SelectNodes('//*'))) { foreach ($a in $fb.Attr) { if ($el.HasAttribute($a)) { $el.RemoveAttribute($a) } } }
        }
        $applied.Add($fb.Name)
        try {
            $r = Import-Xaml $doc.OuterXml
            $State.XamlFallback = $applied -join ', '
            return $r
        } catch { $State.XamlError += "`n`nPo uproszczeniu (" + ($applied -join ', ') + "):`n" + (Get-ExceptionChain $_.Exception) }
    }
    throw $first
}

function Get-ErrText($ErrorRecord) {
    $e = $ErrorRecord.Exception
    while ($e.InnerException) { $e = $e.InnerException }
    return $e.Message
}

function New-Thickness([double]$L, [double]$T, [double]$R, [double]$B) { return (New-Object System.Windows.Thickness $L, $T, $R, $B) }

function New-Brush([string]$Hex, [int]$Alpha = 255) {
    $c = [System.Windows.Media.ColorConverter]::ConvertFromString($Hex)
    $c2 = [System.Windows.Media.Color]::FromArgb([byte]$Alpha, $c.R, $c.G, $c.B)
    $b = New-Object System.Windows.Media.SolidColorBrush $c2
    $b.Freeze()
    return $b
}

function New-TextBlock([string]$Text, [double]$Size, [string]$Color, [string]$Weight = 'Normal', [bool]$Wrap = $false) {
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $Text
    $t.FontSize = $Size
    $t.Foreground = New-Brush $Color
    if ($Weight -eq 'SemiBold') { $t.FontWeight = [System.Windows.FontWeights]::SemiBold }
    elseif ($Weight -eq 'Bold') { $t.FontWeight = [System.Windows.FontWeights]::Bold }
    if ($Wrap) { $t.TextWrapping = [System.Windows.TextWrapping]::Wrap }
    return $t
}

function ConvertTo-Plain([string]$Text) {
    if (-not $Text) { return '' }
    $n = $Text.ToLowerInvariant().Replace('ł', 'l').Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $n.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Get-SafeName([string]$Text) {
    $r = $Text -replace '[\\/:*?"<>|\s]', '_'
    if ($r.Length -gt 60) { $r = $r.Substring(0, 60) }
    return $r
}

function Show-Info([string]$Text) { [void][System.Windows.MessageBox]::Show($Window, $Text, $AppTitle, 'OK', 'Information') }
function Show-Warn([string]$Text) { [void][System.Windows.MessageBox]::Show($Window, $Text, $AppTitle, 'OK', 'Warning') }
function Show-Error([string]$Text) { [void][System.Windows.MessageBox]::Show($Window, $Text, $AppTitle, 'OK', 'Error') }
function Confirm-Action([string]$Text, [string]$Title = $AppTitle) {
    $r = [System.Windows.MessageBox]::Show($Window, $Text, $Title, 'YesNo', 'Question')
    return ($r -eq [System.Windows.MessageBoxResult]::Yes)
}

# ============================================================== log
$LogItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[RegDoktor.LogLine]'

function Add-UiLog([string]$Text, [string]$Kind = 'info') {
    $line = New-Object RegDoktor.LogLine -ArgumentList $Text, $Kind
    $LogItems.Add($line)
    while ($LogItems.Count -gt 4000) { $LogItems.RemoveAt(0) }
    $State.LogBuffer.Add($line.Display)
    $State.LogDirty = $true
}

function Save-LogBuffer {
    if ($State.LogBuffer.Count -eq 0) { return }
    try { [System.IO.File]::AppendAllLines($LogFile, $State.LogBuffer) } catch { }
    $State.LogBuffer.Clear()
}

# ============================================================== background work
function Start-Work {
    param([string]$Title, [scriptblock]$Code, [hashtable]$Vars = @{}, [scriptblock]$OnDone = $null, $Context = $null, [switch]$RestorePoint)
    if ($State.Job) { Show-Info 'Trwa inne zadanie. Poczekaj na jego zakończenie albo kliknij "Anuluj zadanie".'; return $false }
    $Sync.Cancel = $false
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    try { $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass } catch { }
    $rs = [runspacefactory]::CreateRunspace($iss)
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
    $common = @{ Sync = $Sync; OemEncoding = $OemEncoding; BackupRoot = $BackupRoot; ReportRoot = $ReportRoot; WorkRestorePoint = [bool]$RestorePoint; WorkTitle = $Title }
    foreach ($k in $common.Keys) { $rs.SessionStateProxy.SetVariable($k, $common[$k]) }
    foreach ($k in $Vars.Keys) { $rs.SessionStateProxy.SetVariable($k, $Vars[$k]) }
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    $script = $WorkerPrelude + "`n" + 'if ($WorkRestorePoint) { [void](New-RestorePointSafe ("RegDoktor - " + $WorkTitle)) }' + "`n" + $Code.ToString()
    [void]$ps.AddScript($script)
    $State.Job = @{ PS = $ps; RS = $rs; Handle = $ps.BeginInvoke(); Title = $Title; OnDone = $OnDone; Context = $Context; Start = Get-Date; StopRequested = $false }
    Add-UiLog ('▶ ' + $Title) 'head'
    Set-Busy $true $Title
    return $true
}

function Set-Busy([bool]$Busy, [string]$Title = '') {
    $ui.btnCancel.IsEnabled = $Busy
    $ui.pnlActions.IsEnabled = -not $Busy
    foreach ($c in $State.BusyControls) { $c.IsEnabled = -not $Busy }
    if ($Busy) {
        $ui.pbProgress.IsIndeterminate = $true
        $ui.txtStatus.Text = $Title + '…'
    } else {
        $ui.pbProgress.IsIndeterminate = $false
        $ui.pbProgress.Value = 0
        $ui.txtStatus.Text = 'Gotowy.'
        Update-ScanSummary
        $ui.btnHiveCompact.IsEnabled = ($State.Hives.Count -gt 0)
    }
}

function Receive-WorkerMessages([int]$Max = 500) {
    $msg = $null
    $n = 0
    while ($n -lt $Max -and $Sync.Queue.TryDequeue([ref]$msg)) {
        $n++
        if ($msg.T -eq 'log') { Add-UiLog ([string]$msg.Text) ([string]$msg.Kind) }
        elseif ($msg.T -eq 'progress') {
            if ([double]$msg.Value -lt 0) { $ui.pbProgress.IsIndeterminate = $true }
            else {
                $ui.pbProgress.IsIndeterminate = $false
                $ui.pbProgress.Value = [Math]::Max(0, [Math]::Min(100, [double]$msg.Value))
            }
            if ($msg.Text) { $ui.txtStatus.Text = [string]$msg.Text }
        }
    }
}

function Complete-WorkIfDone {
    $j = $State.Job
    if (-not $j) { return }
    if (-not $j.Handle.IsCompleted) {
        if ($State.CancelAt -and -not $j.StopRequested -and ((Get-Date) - $State.CancelAt).TotalSeconds -gt 5) {
            $j.StopRequested = $true
            try { [void]$j.PS.BeginStop($null, $null) } catch { }
        }
        return
    }
    $out = $null
    try { $out = $j.PS.EndInvoke($j.Handle) }
    catch { if (-not $Sync.Cancel) { Add-UiLog ('Zadanie zakończone błędem: ' + (Get-ErrText $_)) 'err' } }
    Receive-WorkerMessages 1000000
    $errs = @($j.PS.Streams.Error)
    $shown = 0
    foreach ($e in $errs) {
        if ($shown -ge 25) { break }
        Add-UiLog ('Błąd: ' + $e.ToString()) 'err'
        $shown++
    }
    if ($errs.Count -gt 25) { Add-UiLog ('… oraz {0} kolejnych błędów' -f ($errs.Count - 25)) 'err' }
    $cancelled = [bool]$Sync.Cancel
    $dur = (Get-Date) - $j.Start
    if ($cancelled) { Add-UiLog ('■ Przerwano: ' + $j.Title) 'warn' }
    else { Add-UiLog ('■ Zakończono: {0}  ({1:hh\:mm\:ss})' -f $j.Title, $dur) 'head' }
    try { $j.PS.Dispose() } catch { }
    try { $j.RS.Close(); $j.RS.Dispose() } catch { }
    $State.Job = $null
    $State.CancelAt = $null
    $Sync.Cancel = $false
    $Sync.Proc = $null
    Set-Busy $false
    if ($j.OnDone) {
        try { & $j.OnDone $out $j.Context $cancelled } catch { Add-UiLog ('Błąd: ' + (Get-ErrText $_)) 'err' }
    }
}

function Get-OutputObjects($Output) {
    $res = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Output) { return ,$res }
    foreach ($o in $Output) { if ($null -ne $o) { $res.Add($o.psobject.BaseObject) } }
    return ,$res
}

$ActionDone = {
    param($Output, $Context, $Cancelled)
    Update-SystemInfo
    if (-not $Cancelled -and $Context -and (@($Context.Tags) -contains 'restart')) {
        if (Confirm-Action ("Zmiany zaczną działać po ponownym uruchomieniu komputera.`n`nZapisz otwarte pliki. Uruchomić ponownie teraz?") 'Wymagany restart') {
            try { Restart-Computer -ErrorAction Stop } catch { Show-Warn ('Nie udało się uruchomić ponownie: ' + (Get-ErrText $_)) }
        }
    }
}

function Invoke-Action($Action) {
    if ($State.Job) { Show-Info 'Trwa inne zadanie. Poczekaj na jego zakończenie albo kliknij "Anuluj zadanie".'; return }
    $confirm = $Action.Confirm
    if ($confirm -is [scriptblock]) { $confirm = [string](& $confirm) }
    if ($confirm) { if (-not (Confirm-Action ([string]$confirm) $Action.Title)) { return } }
    if ($Action.Ui) {
        try { & $Action.Ui } catch { Show-Error (Get-ErrText $_) }
        return
    }
    $rp = [bool]$Action.Risky -and [bool]$ui.chkRestorePoint.IsChecked
    [void](Start-Work -Title $Action.Title -Code $Action.Code -OnDone $ActionDone -Context $Action -RestorePoint:$rp)
}

# ============================================================== system info strip
function Update-SystemInfo {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
        $name = ($os.Caption -replace '^Microsoft\s+', '')
        $ui.txtInfoOS.Text = ('{0} {1}' -f $name, $cv.DisplayVersion).Trim()
        $ui.txtInfoOS.ToolTip = ('Kompilacja {0}.{1}' -f $cv.CurrentBuild, $cv.UBR)
        $up = (Get-Date) - $os.LastBootUpTime
        $ui.txtInfoUptime.Text = ('{0} d {1} h {2} min' -f $up.Days, $up.Hours, $up.Minutes)
        $tot = [double]$os.TotalVisibleMemorySize * 1KB
        $free = [double]$os.FreePhysicalMemory * 1KB
        $pct = 100 * ($tot - $free) / $tot
        $ui.txtInfoRam.Text = ('{0} / {1} ({2:N0}%)' -f (Format-Size ($tot - $free)), (Format-Size $tot), $pct)
        if ($pct -gt 85) { $ui.txtInfoRam.Foreground = New-Brush '#FBBF24' } else { $ui.txtInfoRam.Foreground = New-Brush '#F1F2F5' }
        $dr = New-Object System.IO.DriveInfo $env:SystemDrive
        $ui.txtInfoDisk.Text = ('{0} wolne z {1}' -f (Format-Size $dr.AvailableFreeSpace), (Format-Size $dr.TotalSize))
        if ($dr.AvailableFreeSpace / $dr.TotalSize -lt 0.1) { $ui.txtInfoDisk.Foreground = New-Brush '#FBBF24' } else { $ui.txtInfoDisk.Foreground = New-Brush '#F1F2F5' }
        $ev = Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-Diagnostics-Performance/Operational'; Id = 100 } -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($ev) {
            $bt = 0
            foreach ($d in ([xml]$ev.ToXml()).Event.EventData.Data) { if ($d.Name -eq 'BootTime') { $bt = [double]$d.'#text' } }
            $ui.txtInfoBoot.Text = ('{0:N1} s' -f ($bt / 1000))
            $ui.txtInfoBoot.ToolTip = ('Zmierzono {0:g}' -f $ev.TimeCreated)
        } else { $ui.txtInfoBoot.Text = 'brak danych' }
    } catch { }
}

# ============================================================== repair tiles
function Initialize-ActionTiles {
    $ui.pnlActions.Children.Clear()
    $ui.pnlCatChips.Children.Clear()
    $State.Tiles.Clear()
    $State.Sections.Clear()
    foreach ($catKey in $Categories.Keys) {
        $cat = $Categories[$catKey]
        $acts = @($Actions | Where-Object { $_.Cat -eq $catKey })
        if ($acts.Count -eq 0) { continue }
        $section = New-Object System.Windows.Controls.StackPanel
        $section.Margin = New-Thickness 0 10 0 6
        $hdr = New-Object System.Windows.Controls.StackPanel
        $hdr.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $hdr.Margin = New-Thickness 0 0 0 10
        $dot = New-Object System.Windows.Shapes.Ellipse
        $dot.Width = 12; $dot.Height = 12; $dot.Fill = New-Brush $cat.Color
        $dot.Margin = New-Thickness 0 0 10 0
        $dot.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$hdr.Children.Add($dot)
        [void]$hdr.Children.Add((New-TextBlock $cat.Name 17 '#FFFFFF' 'SemiBold'))
        $sub = New-TextBlock ('   ' + $cat.Desc) 12.5 '#8C93A1'
        $sub.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$hdr.Children.Add($sub)
        [void]$section.Children.Add($hdr)
        $wrap = New-Object System.Windows.Controls.WrapPanel
        [void]$section.Children.Add($wrap)
        foreach ($a in $acts) {
            $btn = New-Object System.Windows.Controls.Button
            $btn.Style = $Window.FindResource('ActionBtn')
            $btn.Background = New-Brush $cat.Color 0x26
            $btn.BorderBrush = New-Brush $cat.Color
            $sp = New-Object System.Windows.Controls.StackPanel
            [void]$sp.Children.Add((New-TextBlock $a.Title 13.5 '#F4F5F7' 'SemiBold' $true))
            $desc = New-TextBlock $a.Desc 12 '#A9AFBC' 'Normal' $true
            $desc.Margin = New-Thickness 0 3 0 0
            [void]$sp.Children.Add($desc)
            $tags = @(@($a.Tags) | Where-Object { $_ })
            if ($a.Risky) { $tags += 'punkt przywracania' }
            if ($tags.Count -gt 0) {
                $tw = New-Object System.Windows.Controls.WrapPanel
                $tw.Margin = New-Thickness 0 7 0 0
                foreach ($tg in $tags) {
                    $chipB = New-Object System.Windows.Controls.Border
                    $chipB.CornerRadius = New-Object System.Windows.CornerRadius 4
                    $chipB.Background = New-Brush $cat.Color 0x33
                    $chipB.Padding = New-Thickness 6 1 6 1
                    $chipB.Margin = New-Thickness 0 0 4 2
                    $chipB.Child = New-TextBlock ([string]$tg) 10.5 $cat.Color 'SemiBold'
                    [void]$tw.Children.Add($chipB)
                }
                [void]$sp.Children.Add($tw)
            }
            $btn.Content = $sp
            $btn.Tag = $a
            $btn.ToolTip = $a.Desc
            $btn.Add_Click({ Invoke-Action $this.Tag })
            [void]$wrap.Children.Add($btn)
            $a.Search = ConvertTo-Plain ($a.Title + ' ' + $a.Desc + ' ' + ($tags -join ' ') + ' ' + $cat.Name)
            $State.Tiles.Add(@{ Button = $btn; Action = $a })
        }
        [void]$ui.pnlActions.Children.Add($section)
        $State.Sections.Add(@{ Panel = $section; Wrap = $wrap })

        $chip = New-Object System.Windows.Controls.Button
        $chip.Style = $Window.FindResource('ChipBtn')
        $chip.Background = New-Brush $cat.Color 0x22
        $chip.BorderBrush = New-Brush $cat.Color 0x70
        $cs = New-Object System.Windows.Controls.StackPanel
        $cs.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $cd = New-Object System.Windows.Shapes.Ellipse
        $cd.Width = 8; $cd.Height = 8; $cd.Fill = New-Brush $cat.Color
        $cd.Margin = New-Thickness 0 0 7 0
        $cd.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$cs.Children.Add($cd)
        [void]$cs.Children.Add((New-TextBlock ('{0}  {1}' -f $cat.Name, $acts.Count) 12.5 '#E6E7EB'))
        $chip.Content = $cs
        $chip.Tag = $section
        $chip.Add_Click({
            try {
                $pt = $this.Tag.TransformToAncestor($ui.pnlActions).Transform((New-Object System.Windows.Point 0, 0))
                $ui.svActions.ScrollToVerticalOffset($pt.Y)
            } catch { $this.Tag.BringIntoView() }
        })
        [void]$ui.pnlCatChips.Children.Add($chip)
    }
}

function Update-ActionFilter {
    $raw = $ui.txtFilter.Text
    if ($raw) { $ui.lblFilterHint.Visibility = 'Collapsed' } else { $ui.lblFilterHint.Visibility = 'Visible' }
    $words = @((ConvertTo-Plain $raw) -split '\s+' | Where-Object { $_ })
    foreach ($t in $State.Tiles) {
        $vis = $true
        foreach ($w in $words) { if (-not $t.Action.Search.Contains($w)) { $vis = $false; break } }
        if ($vis) { $t.Button.Visibility = 'Visible' } else { $t.Button.Visibility = 'Collapsed' }
    }
    foreach ($s in $State.Sections) {
        $any = $false
        foreach ($c in $s.Wrap.Children) { if ($c.Visibility -eq 'Visible') { $any = $true; break } }
        if ($any) { $s.Panel.Visibility = 'Visible' } else { $s.Panel.Visibility = 'Collapsed' }
    }
}

# ============================================================== dialogs
function Set-DarkChrome($Win) {
    $Win.Add_SourceInitialized({
        try { [RegDoktor.Native]::DarkTitleBar((New-Object System.Windows.Interop.WindowInteropHelper $this).Handle) } catch { }
    })
}

function New-AppIcon {
    try {
        $dv = New-Object System.Windows.Media.DrawingVisual
        $dc = $dv.RenderOpen()
        $g = New-Object System.Windows.Media.LinearGradientBrush ([System.Windows.Media.Color]::FromRgb(59, 130, 246)), ([System.Windows.Media.Color]::FromRgb(236, 72, 153)), 45.0
        $dc.DrawRoundedRectangle($g, $null, [System.Windows.Rect]::new(0, 0, 64, 64), 14, 14)
        $ft = New-Object System.Windows.Media.FormattedText 'R', ([System.Globalization.CultureInfo]::InvariantCulture), ([System.Windows.FlowDirection]::LeftToRight), (New-Object System.Windows.Media.Typeface 'Segoe UI Semibold'), 44.0, ([System.Windows.Media.Brushes]::White)
        $dc.DrawText($ft, [System.Windows.Point]::new((64 - $ft.Width) / 2, (64 - $ft.Height) / 2))
        $dc.Close()
        $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap 64, 64, 96.0, 96.0, ([System.Windows.Media.PixelFormats]::Pbgra32)
        $bmp.Render($dv)
        $bmp.Freeze()
        return $bmp
    } catch { return $null }
}

function Join-DialogStyles($Dialog) {
    Set-DarkChrome $Dialog.Window
    $Dialog.Window.Owner = $Window
    try { $Dialog.Window.Resources.MergedDictionaries.Add($Window.Resources) } catch { }
    try { $Dialog.btnOk.Style = $Window.FindResource('PrimaryBtn') } catch { }
}

function Show-InputDialog([string]$Title, [string]$Prompt, [string]$Default = '') {
    $d = Import-Xaml $InputDialogXaml
    Join-DialogStyles $d
    $d.Window.Title = $Title
    $d.txtPrompt.Text = $Prompt
    $d.txtInput.Text = $Default
    $d.btnOk.Add_Click({ [System.Windows.Window]::GetWindow($this).DialogResult = $true })
    $d.Window.Add_ContentRendered({ $t = $this.FindName('txtInput'); [void]$t.Focus(); $t.SelectAll() })
    if ($d.Window.ShowDialog()) { return $d.txtInput.Text.Trim() }
    return $null
}

$ValueHints = @{
    rbSZ    = 'Tekst. Używany najczęściej (ścieżki, nazwy, ustawienia tekstowe).'
    rbEXP   = 'Tekst ze zmiennymi środowiskowymi, np. %SystemRoot%\System32 – rozwijanymi przy odczycie.'
    rbDW    = 'Liczba 32-bitowa: wpisz dziesiętnie (np. 1) lub szesnastkowo z prefiksem 0x (np. 0x1A).'
    rbQW    = 'Liczba 64-bitowa: dziesiętnie lub szesnastkowo z prefiksem 0x.'
    rbMULTI = 'Lista tekstów – każdy element w osobnej linii.'
    rbBIN   = 'Dane binarne jako bajty szesnastkowe, np. 01 a0 ff 00 (spacje i przecinki są ignorowane).'
}

function Show-ValueDialog([string]$Title, [string]$Name, $Kind, [string]$Data, [bool]$IsNew) {
    $kindMap = @{ String = 'rbSZ'; ExpandString = 'rbEXP'; DWord = 'rbDW'; QWord = 'rbQW'; MultiString = 'rbMULTI'; Binary = 'rbBIN'; None = 'rbBIN'; Unknown = 'rbBIN' }
    while ($true) {
        $d = Import-Xaml $ValueDialogXaml
        Join-DialogStyles $d
        $d.Window.Title = $Title
        $d.txtName.Text = $Name
        if (-not $IsNew) { $d.txtName.IsReadOnly = $true; $d.txtName.Opacity = 0.75; $d.pnlTypes.IsEnabled = $false }
        $rb = $kindMap[[string]$Kind]
        if (-not $rb) { $rb = 'rbSZ' }
        $d[$rb].IsChecked = $true
        $d.txtHint.Text = $ValueHints[$rb]
        foreach ($n in @('rbSZ', 'rbEXP', 'rbDW', 'rbQW', 'rbMULTI', 'rbBIN')) {
            $d[$n].Tag = $ValueHints[$n]
            $d[$n].Add_Checked({ $w = [System.Windows.Window]::GetWindow($this); if ($w) { $w.FindName('txtHint').Text = [string]$this.Tag } })
        }
        $d.txtData.Text = $Data
        $d.btnOk.Add_Click({ [System.Windows.Window]::GetWindow($this).DialogResult = $true })
        if ($IsNew) { $d.Window.Add_ContentRendered({ $t = $this.FindName('txtName'); [void]$t.Focus(); $t.SelectAll() }) }
        else { $d.Window.Add_ContentRendered({ $t = $this.FindName('txtData'); [void]$t.Focus(); $t.SelectAll() }) }
        if (-not $d.Window.ShowDialog()) { return $null }
        $sel = [Microsoft.Win32.RegistryValueKind]::String
        if ($d.rbEXP.IsChecked) { $sel = [Microsoft.Win32.RegistryValueKind]::ExpandString }
        elseif ($d.rbDW.IsChecked) { $sel = [Microsoft.Win32.RegistryValueKind]::DWord }
        elseif ($d.rbQW.IsChecked) { $sel = [Microsoft.Win32.RegistryValueKind]::QWord }
        elseif ($d.rbMULTI.IsChecked) { $sel = [Microsoft.Win32.RegistryValueKind]::MultiString }
        elseif ($d.rbBIN.IsChecked) { $sel = [Microsoft.Win32.RegistryValueKind]::Binary }
        if (-not $IsNew) { $sel = [Microsoft.Win32.RegistryValueKind]$Kind }
        $Name = $d.txtName.Text
        $Data = $d.txtData.Text
        $Kind = $sel
        if ($IsNew -and -not $Name.Trim()) { Show-Warn 'Podaj nazwę wartości.'; continue }
        try {
            [void][RegDoktor.Reg]::ParseValue($Data, $sel)
            return @{ Name = $Name; Kind = $sel; Data = $Data }
        } catch { Show-Warn ('Nieprawidłowe dane: ' + (Get-ErrText $_)) }
    }
}

# ============================================================== registry browser
$ProtectedKeys = @(
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft', 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes', 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node',
    'HKEY_LOCAL_MACHINE\SOFTWARE\Policies', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT',
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion',
    'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID',
    'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet', 'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001', 'HKEY_LOCAL_MACHINE\SYSTEM\Setup',
    'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services', 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control', 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum',
    'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services', 'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control', 'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Enum',
    'HKEY_CURRENT_USER\Software\Microsoft', 'HKEY_CURRENT_USER\Software\Classes', 'HKEY_CURRENT_USER\Software\Policies',
    'HKEY_CURRENT_USER\Software\Microsoft\Windows', 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion',
    'HKEY_CLASSES_ROOT\CLSID', 'HKEY_CLASSES_ROOT\Interface', 'HKEY_CLASSES_ROOT\TypeLib', 'HKEY_CLASSES_ROOT\exefile', 'HKEY_CLASSES_ROOT\.exe'
)

function Test-ProtectedKey([string]$Path) {
    $p = [RegDoktor.Reg]::NormalizePath($Path)
    if (-not $p) { return $true }
    if ($p.Split('\').Count -le 2) { return $true }
    if ($p -match '^HKEY_USERS\\[^\\]+\\[^\\]+$') { return $true }
    foreach ($x in $ProtectedKeys) { if ($p -eq $x) { return $true } }
    return $false
}

function Get-CurrentRegPath {
    if ($State.RegPath) { return $State.RegPath }
    $n = $ui.tvReg.SelectedItem
    if ($n -and -not $n.IsDummy) { return $n.FullPath }
    return $null
}

function Show-RegValues([string]$Path) {
    $State.RegPath = $Path
    $ui.txtAddress.Text = $Path
    try { $ui.dgValues.ItemsSource = [RegDoktor.Reg]::GetValues($Path) }
    catch {
        $ui.dgValues.ItemsSource = $null
        Add-UiLog ('Brak dostępu do klucza {0}: {1}' -f $Path, (Get-ErrText $_)) 'warn'
    }
}

function Find-RegNode([string]$Path) {
    $norm = [RegDoktor.Reg]::NormalizePath($Path)
    if (-not $norm) { return $null }
    $parts = $norm.Split('\')
    $node = $null
    foreach ($r in $State.Roots) { if ($r.Name -eq $parts[0]) { $node = $r } }
    for ($i = 1; $i -lt $parts.Length -and $node; $i++) {
        if (-not $node.IsLoaded) { return $null }
        $next = $null
        foreach ($c in $node.Children) { if (-not $c.IsDummy -and $c.Name -eq $parts[$i]) { $next = $c; break } }
        $node = $next
    }
    return $node
}

function Update-RegNode([string]$Path) {
    $n = Find-RegNode $Path
    if ($n) { $n.Reload(); $n.IsExpanded = $true }
}

function Find-ItemsHostPanel($Control) {
    [void]$Control.ApplyTemplate()
    $stack = New-Object System.Collections.Stack
    $stack.Push($Control)
    while ($stack.Count -gt 0) {
        $d = $stack.Pop()
        $cnt = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($d)
        for ($i = 0; $i -lt $cnt; $i++) {
            $c = [System.Windows.Media.VisualTreeHelper]::GetChild($d, $i)
            if ($c -is [System.Windows.Controls.Panel] -and $c.IsItemsHost -and [System.Windows.Controls.ItemsControl]::GetItemsOwner($c) -eq $Control) { return $c }
            if ($c -isnot [System.Windows.Controls.TreeViewItem]) { $stack.Push($c) }
        }
    }
    return $null
}

function Show-PendingNode {
    $chain = $State.PendingChain
    $State.PendingChain = $null
    if (-not $chain) { return }
    try {
        $container = $ui.tvReg
        $tvi = $null
        foreach ($n in $chain) {
            $container.UpdateLayout()
            $tvi = $container.ItemContainerGenerator.ContainerFromItem($n)
            if (-not $tvi) {
                $idx = $container.Items.IndexOf($n)
                $panel = Find-ItemsHostPanel $container
                if ($panel -is [System.Windows.Controls.VirtualizingStackPanel] -and $idx -ge 0) {
                    $panel.BringIndexIntoViewPublic($idx)
                    $container.UpdateLayout()
                    $tvi = $container.ItemContainerGenerator.ContainerFromIndex($idx)
                }
            }
            if (-not $tvi) { return }
            $container = $tvi
        }
        if ($tvi) { $tvi.BringIntoView() }
    } catch { }
}

function Select-RegPath([string]$Path) {
    $norm = [RegDoktor.Reg]::NormalizePath($Path)
    if (-not $norm) { Show-Warn ('Nieprawidłowa ścieżka rejestru: ' + $Path); return }
    $parts = $norm.Split('\')
    $node = $null
    foreach ($r in $State.Roots) { if ($r.Name -eq $parts[0]) { $node = $r } }
    if (-not $node) { return }
    $chain = New-Object System.Collections.Generic.List[object]
    $chain.Add($node)
    for ($i = 1; $i -lt $parts.Length; $i++) {
        $node.IsExpanded = $true
        $child = $node.FindChild($parts[$i])
        if (-not $child) { Add-UiLog ('Nie znaleziono klucza: ' + (($parts[0..$i]) -join '\')) 'warn'; break }
        $node = $child
        $chain.Add($node)
    }
    $prev = $ui.tvReg.SelectedItem
    if ($prev -and $prev -ne $node) { $prev.IsSelected = $false }
    if ($State.SelectedNode -and $State.SelectedNode -ne $node) { $State.SelectedNode.IsSelected = $false }
    $node.IsSelected = $true
    $State.SelectedNode = $node
    Show-RegValues $node.FullPath
    $State.PendingChain = $chain
    [void]$ui.tvReg.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action] { Show-PendingNode })
}

function Backup-RegValues([string]$Path, [string]$Label) {
    $dir = Join-Path $BackupRoot ('Przegladarka_' + (Get-Date -Format 'yyyy-MM-dd'))
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $file = Join-Path $dir ('{0}_{1}_{2}.reg' -f (Get-Date -Format 'HHmmss_fff'), $Label, (Get-SafeName ([RegDoktor.Reg]::Leaf($Path))))
    [void][RegDoktor.Reg]::ExportKeyValues($Path, $file)
    return $file
}

function Invoke-RegNewKey {
    $p = Get-CurrentRegPath
    if (-not $p) { Show-Warn 'Najpierw zaznacz klucz nadrzędny w drzewie.'; return }
    $name = Show-InputDialog 'Nowy klucz' ("Nazwa nowego podklucza w:`n" + $p) 'Nowy klucz #1'
    if (-not $name) { return }
    if ($name.Contains('\')) { Show-Warn 'Nazwa klucza nie może zawierać znaku \'; return }
    try {
        $np = [RegDoktor.Reg]::CreateSubKey($p, $name)
        Add-UiLog ('Utworzono klucz: ' + $np) 'ok'
        Update-RegNode $p
        Select-RegPath $np
    } catch { Show-Error ('Nie udało się utworzyć klucza: ' + (Get-ErrText $_)) }
}

function Invoke-RegNewValue {
    $p = Get-CurrentRegPath
    if (-not $p) { Show-Warn 'Najpierw zaznacz klucz w drzewie.'; return }
    $r = Show-ValueDialog 'Nowa wartość' 'Nowa wartość #1' 'String' '' $true
    if (-not $r) { return }
    try {
        if ([RegDoktor.Reg]::ValueExists($p, $r.Name)) {
            if (-not (Confirm-Action ('Wartość "{0}" już istnieje. Zastąpić ją?' -f $r.Name))) { return }
            $f = Backup-RegValues $p 'nadpisanie'
            Add-UiLog ('Kopia: ' + $f) 'out'
        }
        [RegDoktor.Reg]::SetValue($p, $r.Name, $r.Kind, $r.Data)
        Add-UiLog ('Utworzono wartość {0} ({1}) w {2}' -f $r.Name, [RegDoktor.Reg]::KindName($r.Kind), $p) 'ok'
        Show-RegValues $p
    } catch { Show-Error ('Nie udało się zapisać wartości: ' + (Get-ErrText $_)) }
}

function Invoke-RegEditValue {
    $row = $ui.dgValues.SelectedItem
    if (-not $row) { Show-Warn 'Zaznacz wartość do edycji.'; return }
    $p = $State.RegPath
    $kind = [Microsoft.Win32.RegistryValueKind]::String
    $text = ''
    try { if (-not $row.IsUnset) { $text = [RegDoktor.Reg]::ToEditText($p, $row.Name, [ref]$kind) } }
    catch { Show-Error ('Nie można odczytać wartości: ' + (Get-ErrText $_)); return }
    $r = Show-ValueDialog 'Edycja wartości' $row.Name $kind $text $false
    if (-not $r) { return }
    try {
        $f = Backup-RegValues $p 'edycja'
        [RegDoktor.Reg]::SetValue($p, $row.Name, $r.Kind, $r.Data)
        Add-UiLog ('Zmieniono wartość {0} w {1} (kopia: {2})' -f $row.DisplayName, $p, $f) 'ok'
        Show-RegValues $p
    } catch { Show-Error ('Nie udało się zapisać wartości: ' + (Get-ErrText $_)) }
}

function Invoke-RegDelete([switch]$KeyOnly) {
    $p = Get-CurrentRegPath
    if (-not $p) { return }
    $rows = @()
    if (-not $KeyOnly -and $ui.dgValues.SelectedItems.Count -gt 0) {
        $rows = @($ui.dgValues.SelectedItems | Where-Object { -not $_.IsUnset })
        if ($rows.Count -eq 0) { Show-Info 'Wartość domyślna tego klucza nie jest ustawiona – nie ma czego usuwać.'; return }
    }
    if ($rows.Count -gt 0) {
        $names = ($rows | ForEach-Object { $_.DisplayName }) -join ', '
        if (-not (Confirm-Action ("Usunąć wartości ({0}):`n{1}`n`nz klucza:`n{2}`n`nKopia .reg zostanie zapisana." -f $rows.Count, $names, $p) 'Usuwanie wartości')) { return }
        try { $f = Backup-RegValues $p 'usuniecie' }
        catch { Show-Error ('Nie udało się wykonać kopii – nic nie usunięto. ' + (Get-ErrText $_)); return }
        foreach ($r in $rows) {
            try { [RegDoktor.Reg]::DeleteValue($p, $r.Name); Add-UiLog ('Usunięto wartość {0} z {1}' -f $r.DisplayName, $p) 'ok' }
            catch { Add-UiLog ('Nie usunięto {0}: {1}' -f $r.DisplayName, (Get-ErrText $_)) 'err' }
        }
        Add-UiLog ('Kopia: ' + $f) 'out'
        Show-RegValues $p
        return
    }
    if (Test-ProtectedKey $p) { Show-Warn ("Ten klucz jest niezbędny do działania Windows – usuwanie zostało zablokowane:`n`n" + $p); return }
    if (-not (Confirm-Action ("Usunąć klucz wraz z CAŁĄ zawartością (podklucze i wartości)?`n`n{0}`n`nPrzed usunięciem zostanie zapisana kopia .reg." -f $p) 'Usuwanie klucza')) { return }
    $dir = Join-Path $BackupRoot ('Przegladarka_' + (Get-Date -Format 'yyyy-MM-dd'))
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $file = Join-Path $dir ('{0}_klucz_{1}.reg' -f (Get-Date -Format 'HHmmss_fff'), (Get-SafeName ([RegDoktor.Reg]::Leaf($p))))
    [void](Start-Work -Title ('Usuwanie klucza ' + [RegDoktor.Reg]::Leaf($p)) -Vars @{ KeyPath = $p; BackupFile = $file } -Context $p -Code {
        $rc = Export-RegTree $KeyPath $BackupFile
        if ($rc -ne 0) { Err 'Nie udało się wykonać kopii .reg – klucz NIE został usunięty.'; return }
        try { [RegDoktor.Reg]::DeleteKeyTree($KeyPath); Ok ('Usunięto klucz ' + $KeyPath); Log ('Kopia: ' + $BackupFile) 'out' }
        catch { $m = $_.Exception.Message; if ($_.Exception.InnerException) { $m = $_.Exception.InnerException.Message }; Err ('Nie usunięto klucza: ' + $m) }
    } -OnDone {
        param($o, $ctx)
        $parent = [RegDoktor.Reg]::Parent($ctx)
        Update-RegNode $parent
        Select-RegPath $parent
    })
}

function Invoke-RegExport {
    $p = Get-CurrentRegPath
    if (-not $p) { Show-Warn 'Najpierw zaznacz klucz w drzewie.'; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Pliki rejestru (*.reg)|*.reg'
    $dlg.FileName = (Get-SafeName ([RegDoktor.Reg]::Leaf($p))) + '.reg'
    $dlg.InitialDirectory = $BackupRoot
    if (-not $dlg.ShowDialog($Window)) { return }
    [void](Start-Work -Title ('Eksport ' + $p) -Vars @{ KeyPath = $p; OutFile = $dlg.FileName } -Code {
        $rc = Export-RegTree $KeyPath $OutFile
        if ($rc -eq 0) { Ok ('Wyeksportowano do {0} ({1})' -f $OutFile, (Format-Size (Get-Item -LiteralPath $OutFile).Length)) } else { Err 'Eksport nie powiódł się.' }
    })
}

function Invoke-RegImport([string]$Title = 'Import plików .reg') {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Pliki rejestru (*.reg)|*.reg'
    $dlg.Multiselect = $true
    $dlg.InitialDirectory = $BackupRoot
    $dlg.Title = $Title
    if (-not $dlg.ShowDialog($Window)) { return }
    $files = @($dlg.FileNames)
    $list = ($files | Select-Object -First 12 | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join "`n"
    if ($files.Count -gt 12) { $list += "`n…" }
    if (-not (Confirm-Action ("Zaimportować do rejestru {0} plik(ów)?`n`n{1}" -f $files.Count, $list) $Title)) { return }
    [void](Start-Work -Title $Title -Vars @{ Files = [string[]]$files } -Code {
        $reg = Join-Path $env:SystemRoot 'System32\reg.exe'
        foreach ($f in $Files) {
            $rc = Invoke-Cmd $reg ('import "{0}"' -f $f) -Quiet
            if ($rc -eq 0) { Ok ('Zaimportowano: ' + $f) } else { Err ('Błąd importu: ' + $f) }
        }
    } -OnDone {
        $p = Get-CurrentRegPath
        if ($p) { Update-RegNode $p; Show-RegValues $p }
    })
}

function Restore-RegBackup { Invoke-RegImport 'Przywracanie z kopii .reg' }

function Copy-RegPath {
    $p = Get-CurrentRegPath
    if (-not $p) { return }
    try { [System.Windows.Clipboard]::SetText($p); Add-UiLog ('Skopiowano ścieżkę: ' + $p) 'out' } catch { }
}

function Open-InRegedit {
    $p = Get-CurrentRegPath
    if (-not $p) { return }
    $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit'
    $prefix = $null
    try {
        $last = [string](Get-ItemProperty -LiteralPath $k -Name LastKey -ErrorAction Stop).LastKey
        $i = $last.IndexOf('\HKEY_')
        if ($i -gt 0) { $prefix = $last.Substring(0, $i) } elseif ($last -like 'HKEY_*') { $prefix = '' }
    } catch { }
    if ($null -eq $prefix) { if ((Get-UICulture).Name -like 'pl*') { $prefix = 'Komputer' } else { $prefix = 'Computer' } }
    if ($prefix) { $val = $prefix + '\' + $p } else { $val = $p }
    try {
        if (-not (Test-Path -LiteralPath $k)) { New-Item -Path $k -Force | Out-Null }
        Set-ItemProperty -LiteralPath $k -Name LastKey -Value $val
        Start-Process -FilePath (Join-Path $env:SystemRoot 'regedit.exe') -ArgumentList '/m'
    } catch { Show-Error (Get-ErrText $_) }
}

# ---- search
function Start-RegSearch {
    $t = $ui.txtSearch.Text.Trim()
    if (-not $t) { return }
    if (-not ($ui.chkSKeys.IsChecked -or $ui.chkSNames.IsChecked -or $ui.chkSData.IsChecked)) { Show-Warn 'Zaznacz co najmniej jedno: Klucze, Nazwy wartości lub Dane.'; return }
    if ($ui.rbScopeAll.IsChecked) { $roots = @('HKEY_CURRENT_USER', 'HKEY_LOCAL_MACHINE') }
    else {
        $p = Get-CurrentRegPath
        if (-not $p) { Show-Warn 'Najpierw zaznacz klucz w drzewie albo wybierz "HKCU + HKLM".'; return }
        $roots = @($p)
    }
    $State.SearchResults.Clear()
    $State.Search.Start([string[]]$roots, $t, [bool]$ui.chkSKeys.IsChecked, [bool]$ui.chkSNames.IsChecked, [bool]$ui.chkSData.IsChecked)
    $State.SearchStart = Get-Date
    $State.SearchWasRunning = $true
    $ui.btnSearch.IsEnabled = $false
    $ui.btnSearchStop.IsEnabled = $true
    $ui.txtSearchStatus.Text = 'Szukanie…'
}

function Update-SearchProgress {
    $s = $State.Search
    $hit = $null
    $n = 0
    while ($n -lt 400 -and $s.Results.TryDequeue([ref]$hit)) { $State.SearchResults.Add($hit); $n++ }
    if ($s.IsRunning) {
        $ui.txtSearchStatus.Text = ('Szukanie… przeszukano {0:N0} kluczy, znaleziono {1:N0}   |   {2}' -f $s.Scanned, $s.Found, $s.Current)
    } elseif ($State.SearchWasRunning -and $s.Results.IsEmpty) {
        $State.SearchWasRunning = $false
        $ui.btnSearch.IsEnabled = $true
        $ui.btnSearchStop.IsEnabled = $false
        $msg = 'Zakończono: przeszukano {0:N0} kluczy, znaleziono {1:N0} ({2:N1} s).' -f $s.Scanned, $s.Found, ((Get-Date) - $State.SearchStart).TotalSeconds
        if ($s.Found -ge $s.MaxResults) { $msg += ' Osiągnięto limit wyników – zawęź wyszukiwanie.' }
        $ui.txtSearchStatus.Text = $msg + '  Dwuklik na wyniku przenosi do klucza.'
    }
}

function Open-SearchHit {
    $h = $ui.dgSearch.SelectedItem
    if (-not $h) { return }
    Select-RegPath $h.Path
    if ($h.ValueName) {
        foreach ($row in @($ui.dgValues.ItemsSource)) {
            if ($row.DisplayName -eq $h.ValueName -or $row.Name -eq $h.ValueName) { $ui.dgValues.SelectedItem = $row; $ui.dgValues.ScrollIntoView($row); break }
        }
    }
}

# ============================================================== orphans
function Initialize-ScanCategories {
    $riskColor = @{ 1 = '#34D399'; 2 = '#FBBF24'; 3 = '#F87171' }
    foreach ($c in $ScanCategories) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.IsChecked = ($c.Risk -lt 3)
        $cb.Tag = $c.Id
        $cb.Margin = New-Thickness 0 0 0 10
        $cb.VerticalContentAlignment = [System.Windows.VerticalAlignment]::Top
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Margin = New-Thickness 4 -2 0 0
        $row = New-Object System.Windows.Controls.StackPanel
        $row.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $dot = New-Object System.Windows.Shapes.Ellipse
        $dot.Width = 8; $dot.Height = 8; $dot.Fill = New-Brush $riskColor[[int]$c.Risk]
        $dot.Margin = New-Thickness 0 0 7 0
        $dot.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$row.Children.Add($dot)
        [void]$row.Children.Add((New-TextBlock $c.Name 13 '#E6E7EB' 'SemiBold'))
        [void]$sp.Children.Add($row)
        $ds = New-TextBlock $c.Desc 11.5 '#8C93A1' 'Normal' $true
        $ds.MaxWidth = 260
        [void]$sp.Children.Add($ds)
        $cb.Content = $sp
        [void]$ui.pnlScanCats.Children.Add($cb)
    }
}

function Update-ScanSummary {
    $all = $State.CleanItems.Count
    if ($all -eq 0) {
        $ui.btnClean.IsEnabled = $false
        return
    }
    $sel = 0; $lo = 0; $mid = 0; $hi = 0
    foreach ($i in $State.CleanItems) {
        if ($i.Selected) { $sel++ }
        switch ($i.RiskLevel) { 1 { $lo++ } 2 { $mid++ } 3 { $hi++ } }
    }
    $ui.txtScanSummary.Text = ('Znaleziono {0}  (niskie {1}, średnie {2}, wysokie {3})   •   zaznaczono {4}' -f $all, $lo, $mid, $hi, $sel)
    $ui.btnClean.IsEnabled = ($sel -gt 0 -and -not $State.Job)
}

function Start-OrphanScan {
    if ($State.Job) { Show-Info 'Trwa inne zadanie. Poczekaj na jego zakończenie.'; return }
    $cats = @($ui.pnlScanCats.Children | Where-Object { $_.IsChecked } | ForEach-Object { [string]$_.Tag })
    if ($cats.Count -eq 0) { Show-Warn 'Zaznacz co najmniej jedną kategorię skanowania.'; return }
    $State.CleanItems.Clear()
    $ui.txtScanSummary.Text = 'Skanowanie…'
    [void](Start-Work -Title 'Skanowanie osieroconych wpisów rejestru' -Code $OrphanScanCode -Vars @{ Cats = [string[]]$cats } -OnDone {
        param($out)
        foreach ($o in (Get-OutputObjects $out)) { if ($o -is [RegDoktor.CleanItem]) { $State.CleanItems.Add($o) } }
        if ($State.CleanItems.Count -eq 0) { $ui.txtScanSummary.Text = 'Nie znaleziono osieroconych wpisów.' }
        Update-ScanSummary
    })
}

function Start-OrphanClean {
    $items = @($State.CleanItems | Where-Object { $_.Selected })
    if ($items.Count -eq 0) { return }
    $hi = @($items | Where-Object { $_.RiskLevel -ge 3 }).Count
    $msg = "Usunąć {0} zaznaczonych wpisów z rejestru?`n`nKażdy wpis zostanie najpierw zapisany w kopii .reg – możesz go przywrócić przyciskiem 'Przywróć z kopii…'." -f $items.Count
    if ($hi -gt 0) { $msg += ("`n`nUWAGA: zaznaczono {0} wpisów wysokiego ryzyka (usługi/sterowniki)." -f $hi) }
    if (-not (Confirm-Action $msg 'Czyszczenie rejestru')) { return }
    [void](Start-Work -Title ('Usuwanie {0} osieroconych wpisów' -f $items.Count) -Code $OrphanCleanCode -Vars @{ Items = [object[]]$items } -RestorePoint:([bool]$ui.chkRestorePoint.IsChecked) -OnDone {
        param($out)
        foreach ($o in (Get-OutputObjects $out)) { if ($o -is [RegDoktor.CleanItem]) { [void]$State.CleanItems.Remove($o) } }
        Update-ScanSummary
        if ($State.CleanItems.Count -eq 0) { $ui.txtScanSummary.Text = 'Wszystkie zaznaczone wpisy zostały usunięte.' }
    })
}

# ============================================================== hive compaction
function Start-HiveAnalyze {
    if ($State.Job) { Show-Info 'Trwa inne zadanie. Poczekaj na jego zakończenie.'; return }
    $State.Hives.Clear()
    $ui.txtHiveSummary.Text = 'Analiza…'
    [void](Start-Work -Title 'Analiza fragmentacji rejestru' -Code $HiveAnalyzeCode -OnDone {
        param($out)
        $tot = 0; $save = 0
        foreach ($o in (Get-OutputObjects $out)) {
            if ($o -is [RegDoktor.HiveInfo]) {
                $State.Hives.Add($o)
                if ($o.Compacted -gt 0) { $tot += $o.Size; $save += ($o.Size - $o.Compacted) }
            }
        }
        if ($tot -gt 0) { $ui.txtHiveSummary.Text = ('Rozmiar gałęzi: {0}  •  możliwa oszczędność: {1} ({2:N1}%). Zaznaczono gałęzie z zyskiem ≥ 5%.' -f (Format-Size $tot), (Format-Size $save), (100.0 * $save / $tot)) }
        else { $ui.txtHiveSummary.Text = 'Brak danych analizy.' }
        $ui.btnHiveCompact.IsEnabled = ($State.Hives.Count -gt 0)
    })
}

function Start-HiveCompact {
    $sel = @($State.Hives | Where-Object { $_.Selected -and $_.Compacted -gt 0 })
    if ($sel.Count -eq 0) { Show-Warn 'Zaznacz gałęzie do kompaktowania (najpierw wykonaj analizę).'; return }
    $msg = "Wybrane gałęzie ({0}) zostaną podmienione na skompaktowane kopie przy NAJBLIŻSZYM RESTARCIE.`n`n• Oryginalne pliki zostaną zachowane jako .old wraz z instrukcją przywrócenia.`n• Uruchom komputer ponownie zaraz po zakończeniu – zmiany ustawień wprowadzone przed restartem mogą zostać utracone.`n`nKontynuować?" -f $sel.Count
    if (-not (Confirm-Action $msg 'Kompaktowanie rejestru')) { return }
    [void](Start-Work -Title 'Kompaktowanie rejestru' -Code $HiveCompactCode -Vars @{ Hives = [object[]]$sel } -RestorePoint:([bool]$ui.chkRestorePoint.IsChecked) -OnDone {
        param($out, $ctx, $cancelled)
        $r = $null
        foreach ($o in (Get-OutputObjects $out)) { if ($o -is [hashtable]) { $r = $o } }
        foreach ($h in $State.Hives) { if ($h.Selected) { $h.Status = 'Zaplanowano (po restarcie)' } }
        if ($r -and $r.Scheduled -gt 0) {
            $ui.txtHiveSummary.Text = ('Zaplanowano kompaktowanie {0} gałęzi – wymagany restart.' -f $r.Scheduled)
            if (Confirm-Action ("Kompaktowanie zostanie dokończone podczas ponownego uruchomienia.`n`nZapisz otwarte pliki. Uruchomić ponownie teraz?") 'Wymagany restart') {
                try { Restart-Computer -ErrorAction Stop } catch { Show-Warn ('Nie udało się uruchomić ponownie: ' + (Get-ErrText $_)) }
            }
        }
    })
}

# ============================================================== startup manager
function Test-StartupEnabled([string]$ApprovedKey, [string]$Name) {
    $k = [RegDoktor.Reg]::TryOpen($ApprovedKey, $false)
    if (-not $k) { return $true }
    try {
        $v = $k.GetValue($Name)
        if ($v -is [byte[]] -and $v.Length -gt 0) { return (($v[0] -band 1) -eq 0) }
        return $true
    } finally { $k.Close() }
}

function Get-StartupItems {
    $list = New-Object 'System.Collections.Generic.List[RegDoktor.StartupItem]'
    $ap = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\'
    $defs = @(
        @{ Key = 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run'; Appr = 'HKEY_CURRENT_USER\' + $ap + 'Run'; Scope = 'Bieżący użytkownik' },
        @{ Key = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'; Appr = 'HKEY_LOCAL_MACHINE\' + $ap + 'Run'; Scope = 'Wszyscy użytkownicy' },
        @{ Key = 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Appr = 'HKEY_LOCAL_MACHINE\' + $ap + 'Run32'; Scope = 'Wszyscy (32-bit)' }
    )
    foreach ($d in $defs) {
        $k = [RegDoktor.Reg]::TryOpen($d.Key, $false)
        if (-not $k) { continue }
        try {
            foreach ($n in $k.GetValueNames()) {
                if (-not $n) { continue }
                $cmd = [string]$k.GetValue($n, '', [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $it = New-Object RegDoktor.StartupItem
                $it.Name = $n; $it.Command = $cmd; $it.Location = $d.Key; $it.Scope = $d.Scope
                $it.RunKeyPath = $d.Key; $it.ApprovedKeyPath = $d.Appr
                $it.Enabled = Test-StartupEnabled $d.Appr $n
                $it.Missing = [bool](Get-MissingTarget $cmd)
                $list.Add($it)
            }
        } finally { $k.Close() }
    }
    $shell = $null
    try { $shell = New-Object -ComObject WScript.Shell } catch { }
    $folders = @(
        @{ Path = [Environment]::GetFolderPath('Startup'); Appr = 'HKEY_CURRENT_USER\' + $ap + 'StartupFolder'; Scope = 'Bieżący użytkownik (folder)' },
        @{ Path = [Environment]::GetFolderPath('CommonStartup'); Appr = 'HKEY_LOCAL_MACHINE\' + $ap + 'StartupFolder'; Scope = 'Wszyscy (folder)' }
    )
    foreach ($f in $folders) {
        if (-not $f.Path -or -not (Test-Path -LiteralPath $f.Path)) { continue }
        foreach ($file in @(Get-ChildItem -LiteralPath $f.Path -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'desktop.ini' })) {
            $it = New-Object RegDoktor.StartupItem
            $it.Name = $file.BaseName; $it.FilePath = $file.FullName; $it.Location = $f.Path; $it.Scope = $f.Scope
            $it.ApprovedKeyPath = $f.Appr
            $it.Command = $file.FullName
            if ($file.Extension -eq '.lnk' -and $shell) {
                try {
                    $l = $shell.CreateShortcut($file.FullName)
                    $it.Command = ('"{0}" {1}' -f $l.TargetPath, $l.Arguments).Trim()
                    if ($l.TargetPath) { $it.Missing = [bool](Get-MissingTarget ('"' + $l.TargetPath + '"')) }
                } catch { }
            }
            $it.Enabled = Test-StartupEnabled $f.Appr $file.Name
            $list.Add($it)
        }
    }
    return ,$list
}

function Update-StartupList {
    try {
        $items = Get-StartupItems
        $ui.dgStartup.ItemsSource = $items
        $on = @($items | Where-Object { $_.Enabled }).Count
        $miss = @($items | Where-Object { $_.Missing }).Count
        $txt = 'Pozycji: {0}, włączonych: {1}.' -f $items.Count, $on
        if ($miss) { $txt += (' Wskazujących na brakujący plik: {0} (można je usunąć).' -f $miss) }
        $ui.txtStartSummary.Text = $txt + ' Wyłączanie działa jak w Menedżerze zadań i można je cofnąć.'
    } catch { Add-UiLog ('Autostart: ' + (Get-ErrText $_)) 'err' }
}

function Get-StartupValueName($Item) { if ($Item.FilePath) { return [System.IO.Path]::GetFileName($Item.FilePath) } return $Item.Name }

function Set-StartupEnabled([bool]$Enable) {
    $sel = @($ui.dgStartup.SelectedItems)
    if ($sel.Count -eq 0) { Show-Warn 'Zaznacz pozycje autostartu.'; return }
    foreach ($it in $sel) {
        try {
            $p = 'Registry::' + $it.ApprovedKeyPath
            if (-not (Test-Path -LiteralPath $p)) { New-Item -Path $p -Force | Out-Null }
            if ($Enable) { $bytes = [byte[]](2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) }
            else { $bytes = [byte[]](@([byte]3, [byte]0, [byte]0, [byte]0) + [BitConverter]::GetBytes([DateTime]::Now.ToFileTime())) }
            New-ItemProperty -LiteralPath $p -Name (Get-StartupValueName $it) -Value $bytes -PropertyType Binary -Force | Out-Null
            if ($Enable) { Add-UiLog ('Autostart włączony: ' + $it.Name) 'ok' } else { Add-UiLog ('Autostart wyłączony: ' + $it.Name) 'ok' }
        } catch { Add-UiLog ('{0}: {1}' -f $it.Name, (Get-ErrText $_)) 'err' }
    }
    Update-StartupList
}

function Remove-StartupItem {
    $sel = @($ui.dgStartup.SelectedItems)
    if ($sel.Count -eq 0) { Show-Warn 'Zaznacz pozycje autostartu.'; return }
    if (-not (Confirm-Action ("Usunąć z autostartu {0} pozycji?`n`n{1}`n`nWpisy rejestru trafią do kopii .reg, a pliki z folderu Autostart – do folderu kopii." -f $sel.Count, (($sel | ForEach-Object { $_.Name }) -join "`n")) 'Autostart')) { return }
    $dir = Join-Path $BackupRoot ('Autostart_' + (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    foreach ($it in $sel) {
        try {
            if ($it.RunKeyPath) {
                [void][RegDoktor.Reg]::ExportKeyValues($it.RunKeyPath, (Join-Path $dir ('{0}.reg' -f (Get-SafeName $it.Name))))
                [RegDoktor.Reg]::DeleteValue($it.RunKeyPath, $it.Name)
            } elseif ($it.FilePath) {
                Move-Item -LiteralPath $it.FilePath -Destination $dir -Force -ErrorAction Stop
            }
            $vn = Get-StartupValueName $it
            if ([RegDoktor.Reg]::ValueExists($it.ApprovedKeyPath, $vn)) { [RegDoktor.Reg]::DeleteValue($it.ApprovedKeyPath, $vn) }
            Add-UiLog ('Usunięto z autostartu: ' + $it.Name) 'ok'
        } catch { Add-UiLog ('{0}: {1}' -f $it.Name, (Get-ErrText $_)) 'err' }
    }
    Add-UiLog ('Kopia: ' + $dir) 'out'
    Update-StartupList
}

function Open-StartupLocation {
    $it = $ui.dgStartup.SelectedItem
    if (-not $it) { return }
    $target = $it.FilePath
    if (-not $target) { $target = Get-CmdPath $it.Command }
    if (-not $target) { Select-RegPath $it.RunKeyPath; $ui.tabMain.SelectedItem = $ui.tabReg; return }
    if ([System.IO.File]::Exists($target)) { Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList ('/select,"{0}"' -f $target) }
    else { Show-Warn ('Plik nie istnieje: ' + $target) }
}

# ============================================================== build window
try {
    $main = Import-MainXaml $MainXaml
    $Window = $main.Window
    $ui = $main

    $Window.Title = '{0} {1} — rejestr, naprawy i optymalizacja Windows' -f $AppTitle, $AppVersion
    Set-DarkChrome $Window
    $icon = New-AppIcon
    if ($icon) { $Window.Icon = $icon }
    $Window.Dispatcher.Add_UnhandledException({
        param($s, $e)
        try { Add-UiLog ('Nieobsłużony błąd: ' + $e.Exception.Message) 'err' } catch { }
        $e.Handled = $true
    })

    $ui.lstLog.ItemsSource = $LogItems
    $State.CleanItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[RegDoktor.CleanItem]'
    $ui.dgClean.ItemsSource = $State.CleanItems
    $State.Hives = New-Object 'System.Collections.ObjectModel.ObservableCollection[RegDoktor.HiveInfo]'
    $ui.dgHives.ItemsSource = $State.Hives
    $State.Roots = [RegDoktor.RegNode]::CreateRoots()
    $ui.tvReg.ItemsSource = $State.Roots
    $State.Search = New-Object RegDoktor.RegSearch
    $State.SearchResults = New-Object 'System.Collections.ObjectModel.ObservableCollection[RegDoktor.SearchHit]'
    $ui.dgSearch.ItemsSource = $State.SearchResults
    $State.BusyControls = @($ui.btnScan, $ui.btnHiveAnalyze, $ui.btnRestorePoint, $ui.btnRestoreBackup, $ui.btnExportReg, $ui.btnImportReg, $ui.btnClean, $ui.btnHiveCompact)

    $ui.txtAdmin.Text = 'Administrator  •  PowerShell {0}' -f $PSVersionTable.PSVersion.ToString(2)

    Initialize-ActionTiles
    Initialize-ScanCategories

    # ---- header
    $ui.btnRestorePoint.Add_Click({ [void](Start-Work -Title 'Punkt przywracania systemu' -Code { [void](New-RestorePointSafe 'RegDoktor - punkt reczny') }) })
    $ui.btnBackups.Add_Click({ Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList ('"{0}"' -f $AppDataRoot) })

    # ---- repairs tab
    $ui.txtFilter.Add_TextChanged({ Update-ActionFilter })
    $ui.btnInfoRefresh.Add_Click({ Update-SystemInfo })

    # ---- registry browser
    $ui.tvReg.Add_SelectedItemChanged({
        $n = $ui.tvReg.SelectedItem
        if ($n -and -not $n.IsDummy) {
            $State.SelectedNode = $n
            if ($n.FullPath -ne $State.RegPath) { Show-RegValues $n.FullPath }
        }
    })
    $ui.tvReg.Add_PreviewMouseRightButtonDown({
        try {
            $d = $_.OriginalSource
            while ($d -and -not ($d -is [System.Windows.Controls.TreeViewItem])) { $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d) }
            if ($d) { $d.IsSelected = $true; [void]$d.Focus() }
        } catch { }
    })
    $ui.tvReg.Add_PreviewKeyDown({
        if ($_.Key -eq 'F5') { $p = Get-CurrentRegPath; if ($p) { Update-RegNode $p; Show-RegValues $p }; $_.Handled = $true }
        elseif ($_.Key -eq 'Delete') { Invoke-RegDelete -KeyOnly; $_.Handled = $true }
    })
    $ui.txtAddress.Add_KeyDown({ if ($_.Key -eq 'Return') { Select-RegPath $ui.txtAddress.Text; $_.Handled = $true } })
    $ui.btnGo.Add_Click({ Select-RegPath $ui.txtAddress.Text })
    $ui.btnUp.Add_Click({ $p = Get-CurrentRegPath; if ($p) { $par = [RegDoktor.Reg]::Parent($p); if ($par) { Select-RegPath $par } } })
    $ui.btnRegRefresh.Add_Click({ $p = Get-CurrentRegPath; if ($p) { Update-RegNode $p; Show-RegValues $p } })
    $ui.btnNewKey.Add_Click({ Invoke-RegNewKey })
    $ui.btnNewValue.Add_Click({ Invoke-RegNewValue })
    $ui.btnEditValue.Add_Click({ Invoke-RegEditValue })
    $ui.btnDeleteReg.Add_Click({ Invoke-RegDelete })
    $ui.btnExportReg.Add_Click({ Invoke-RegExport })
    $ui.btnImportReg.Add_Click({ Invoke-RegImport })
    $ui.btnCopyPath.Add_Click({ Copy-RegPath })
    $ui.btnOpenRegedit.Add_Click({ Open-InRegedit })
    $ui.dgValues.Add_MouseDoubleClick({ if ($ui.dgValues.SelectedItem) { Invoke-RegEditValue } })
    $ui.dgValues.Add_PreviewKeyDown({
        if ($_.Key -eq 'Delete') { Invoke-RegDelete; $_.Handled = $true }
        elseif ($_.Key -eq 'Return') { Invoke-RegEditValue; $_.Handled = $true }
        elseif ($_.Key -eq 'F5') { $p = Get-CurrentRegPath; if ($p) { Show-RegValues $p }; $_.Handled = $true }
    })

    $cmTree = New-Object System.Windows.Controls.ContextMenu
    foreach ($mi in @(
            @('Nowy klucz', { Invoke-RegNewKey }), @('Nowa wartość', { Invoke-RegNewValue }), @('-', $null),
            @('Usuń klucz', { Invoke-RegDelete -KeyOnly }), @('Eksportuj do .reg', { Invoke-RegExport }), @('-', $null),
            @('Kopiuj ścieżkę', { Copy-RegPath }), @('Otwórz w Regedit', { Open-InRegedit }), @('Odśwież', { $p = Get-CurrentRegPath; if ($p) { Update-RegNode $p; Show-RegValues $p } }))) {
        if ($mi[0] -eq '-') { [void]$cmTree.Items.Add((New-Object System.Windows.Controls.Separator)); continue }
        $m = New-Object System.Windows.Controls.MenuItem
        $m.Header = $mi[0]
        $m.Add_Click($mi[1])
        [void]$cmTree.Items.Add($m)
    }
    $ui.tvReg.ContextMenu = $cmTree
    $cmVal = New-Object System.Windows.Controls.ContextMenu
    foreach ($mi in @(
            @('Edytuj…', { Invoke-RegEditValue }), @('Nowa wartość…', { Invoke-RegNewValue }), @('Usuń', { Invoke-RegDelete }), @('-', $null),
            @('Kopiuj nazwę', { $r = $ui.dgValues.SelectedItem; if ($r) { try { [System.Windows.Clipboard]::SetText($r.Name) } catch { } } }),
            @('Kopiuj dane', { $r = $ui.dgValues.SelectedItem; if ($r) { try { [System.Windows.Clipboard]::SetText([string]$r.Data) } catch { } } }))) {
        if ($mi[0] -eq '-') { [void]$cmVal.Items.Add((New-Object System.Windows.Controls.Separator)); continue }
        $m = New-Object System.Windows.Controls.MenuItem
        $m.Header = $mi[0]
        $m.Add_Click($mi[1])
        [void]$cmVal.Items.Add($m)
    }
    $ui.dgValues.ContextMenu = $cmVal

    $ui.btnSearch.Add_Click({ Start-RegSearch })
    $ui.txtSearch.Add_KeyDown({ if ($_.Key -eq 'Return') { Start-RegSearch; $_.Handled = $true } })
    $ui.btnSearchStop.Add_Click({ $State.Search.Cancel() })
    $ui.dgSearch.Add_MouseDoubleClick({ Open-SearchHit })

    # ---- orphans
    $ui.btnScan.Add_Click({ Start-OrphanScan })
    $ui.btnClean.Add_Click({ Start-OrphanClean })
    $ui.btnRestoreBackup.Add_Click({ Restore-RegBackup })
    $ui.btnSelSafe.Add_Click({ foreach ($i in $State.CleanItems) { $i.Selected = ($i.RiskLevel -eq 1) }; Update-ScanSummary })
    $ui.btnSelAll.Add_Click({ foreach ($i in $State.CleanItems) { $i.Selected = $true }; Update-ScanSummary })
    $ui.btnSelNone.Add_Click({ foreach ($i in $State.CleanItems) { $i.Selected = $false }; Update-ScanSummary })
    $ui.dgClean.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler] { Update-ScanSummary })
    $ui.dgClean.Add_MouseDoubleClick({
        $it = $ui.dgClean.SelectedItem
        if ($it) { $ui.tabMain.SelectedItem = $ui.tabReg; Select-RegPath $it.KeyPath }
    })

    # ---- hives
    $ui.btnHiveAnalyze.Add_Click({ Start-HiveAnalyze })
    $ui.btnHiveCompact.Add_Click({ Start-HiveCompact })

    # ---- startup
    $ui.btnStartRefresh.Add_Click({ Update-StartupList })
    $ui.btnStartEnable.Add_Click({ Set-StartupEnabled $true })
    $ui.btnStartDisable.Add_Click({ Set-StartupEnabled $false })
    $ui.btnStartDelete.Add_Click({ Remove-StartupItem })
    $ui.btnStartOpen.Add_Click({ Open-StartupLocation })
    $ui.dgStartup.Add_MouseDoubleClick({ Open-StartupLocation })

    # ---- log panel
    $ui.btnCancel.Add_Click({
        if (-not $State.Job) { return }
        if (-not (Confirm-Action 'Przerwać bieżące zadanie?' 'Anulowanie')) { return }
        $Sync.Cancel = $true
        $State.CancelAt = Get-Date
        if ($Sync.Proc) { try { $Sync.Proc.Kill() } catch { } }
        Add-UiLog 'Anulowanie zadania…' 'warn'
    })
    $ui.btnLogCopy.Add_Click({
        $sel = @($ui.lstLog.SelectedItems)
        if ($sel.Count -eq 0) { $sel = @($LogItems) }
        try { [System.Windows.Clipboard]::SetText((($sel | ForEach-Object { $_.Display }) -join "`r`n")) } catch { }
    })
    $ui.btnLogSave.Add_Click({
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = 'Plik tekstowy (*.txt)|*.txt'
        $dlg.FileName = 'RegDoktor_dziennik_{0}.txt' -f (Get-Date -Format 'yyyy-MM-dd_HH-mm')
        $dlg.InitialDirectory = $ReportRoot
        if ($dlg.ShowDialog($Window)) { [System.IO.File]::WriteAllLines($dlg.FileName, [string[]]@($LogItems | ForEach-Object { $_.Display }), [System.Text.Encoding]::UTF8) }
    })
    $ui.btnLogClear.Add_Click({ $LogItems.Clear() })

    # ---- timer: worker messages, job completion, search results, log file
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(120)
    $timer.Add_Tick({
        try { Receive-WorkerMessages } catch { }
        try { Complete-WorkIfDone } catch { Add-UiLog ('Błąd: ' + (Get-ErrText $_)) 'err' }
        try { Update-SearchProgress } catch { }
        if ($State.LogDirty) {
            $State.LogDirty = $false
            if ($LogItems.Count -gt 0) { $ui.lstLog.ScrollIntoView($LogItems[$LogItems.Count - 1]) }
            Save-LogBuffer
        }
    })

    $Window.Add_ContentRendered({
        Update-SystemInfo
        Update-StartupList
        Select-RegPath 'HKEY_CURRENT_USER\Software'
        $timer.Start()
    })
    $Window.Add_Closing({
        param($s, $e)
        if ($State.Job) {
            if (-not (Confirm-Action 'Trwa zadanie. Przerwać je i zamknąć program?' 'Zamykanie')) { $e.Cancel = $true; return }
            $Sync.Cancel = $true
            if ($Sync.Proc) { try { $Sync.Proc.Kill() } catch { } }
            try { $State.Job.PS.Stop() } catch { }
        }
        try { $State.Search.Cancel() } catch { }
        $timer.Stop()
        Save-LogBuffer
    })

    Add-UiLog ('{0} {1} uruchomiony. Kopie zapasowe i raporty: {2}' -f $AppTitle, $AppVersion, $AppDataRoot) 'head'
    Add-UiLog 'Wskazówka: najedź na kafelek, aby zobaczyć opis. Kafelki z etykietą "punkt przywracania" tworzą go automatycznie przed zmianami.'
    if ($State.XamlFallback) {
        Add-UiLog ('Okno uruchomiono w trybie zgodności (wyłączono: {0}). Wszystkie funkcje działają, zmienił się tylko wygląd.' -f $State.XamlFallback) 'warn'
        foreach ($l in ($State.XamlError -split "`n")) { if ($l.Trim()) { Add-UiLog ('   ' + $l) 'out' } }
        try { Set-Content -LiteralPath (Join-Path $AppDataRoot 'diagnostyka_okna.txt') -Value ("Tryb zgodności: " + $State.XamlFallback + "`r`n`r`n" + $State.XamlError) -Encoding UTF8 } catch { }
    }

    [void]$Window.ShowDialog()
} catch {
    $err = $_
    $msg = "Wystąpił błąd uruchomienia:`n`n" + (Get-ExceptionChain $err.Exception)
    if ($err.InvocationInfo) { $msg += "`n`nLinia skryptu: " + $err.InvocationInfo.ScriptLineNumber }
    if ($State.XamlError) { $msg += "`n`nSzczegóły ładowania okna:`n" + $State.XamlError }
    $diag = Join-Path $AppDataRoot 'blad_uruchomienia.txt'
    try {
        $full = $msg + "`r`n`r`n" + $err.ScriptStackTrace + "`r`n`r`n" + $err.Exception.ToString() +
            ("`r`n`r`nPowerShell {0}, .NET {1}, Windows {2}" -f $PSVersionTable.PSVersion, [Environment]::Version, [Environment]::OSVersion.Version)
        Set-Content -LiteralPath $diag -Value $full -Encoding UTF8
        Add-Content -LiteralPath $LogFile -Value ('[FATAL] ' + $msg) -Encoding UTF8
    } catch { }
    if ($msg.Length -gt 1800) { $msg = $msg.Substring(0, 1800) + ' …' }
    Show-StartupError ($msg + "`n`nPełny opis zapisano w pliku:`n" + $diag + "`n(Ctrl+C w tym oknie kopiuje jego treść.)")
}
Save-LogBuffer
[Environment]::Exit(0)
