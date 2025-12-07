using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using Timer = System.Windows.Forms.Timer;

static class Program
{
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr childAfter, string cls, string title);

    private static bool IsTrayReady()
    {
        IntPtr taskbar = FindWindow("Shell_TrayWnd", null);
        if (taskbar == IntPtr.Zero) return false;
        IntPtr tray = FindWindowEx(taskbar, IntPtr.Zero, "TrayNotifyWnd", null);
        return tray != IntPtr.Zero;
    }

    private static void WaitForExplorerAndTray(int timeoutMs)
    {
        int start = Environment.TickCount;
        while (Process.GetProcessesByName("explorer").Length == 0)
        {
            if (Environment.TickCount - start > timeoutMs) break;
            Thread.Sleep(300);
        }
        while (!IsTrayReady())
        {
            if (Environment.TickCount - start > timeoutMs) break;
            Thread.Sleep(300);
        }
    }

    [STAThread]
    static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Global shutdown guards
        SystemEvents.SessionEnding += (s, e) => TrayContext.BeginShutdown("Program.SessionEnding");
        SystemEvents.SessionEnded += (s, e) => TrayContext.BeginShutdown("Program.SessionEnded");
        AppDomain.CurrentDomain.ProcessExit += (s, e) => TrayContext.BeginShutdown("Program.ProcessExit");

        Application.ThreadException += (s, e) => TrayContext.LogAndShow("ThreadException", e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            TrayContext.LogAndShow("UnhandledException", e.ExceptionObject as Exception);

        try
        {
            // TrayContext.EnsureIdentityAndShortcut();
            WaitForExplorerAndTray(7000);
            bool debug = (args.Length > 0 && string.Equals(args[0], "/debug", StringComparison.OrdinalIgnoreCase));
            Application.Run(new TrayContext(debug));
        }
        catch (Exception ex)
        {
            TrayContext.LogAndShow("MainCatch", ex);
        }
    }
}

public sealed class TrayContext : ApplicationContext
{
    // === PowrProf interop (no external powercfg.exe) ===
    private const uint ACCESS_SCHEME = 16; // enumerate power schemes
    private const uint ERROR_NO_MORE_ITEMS = 259;

    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerEnumerate(
        IntPtr RootPowerKey, IntPtr SchemeGuid, IntPtr SubGroupOfPowerSettingsGuid,
        uint AccessFlags, uint Index, IntPtr Buffer, ref uint BufferSize);

    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerReadFriendlyName(
        IntPtr RootPowerKey, ref Guid SchemeGuid, IntPtr SubGroupOfPowerSettingsGuid,
        IntPtr PowerSettingGuid, IntPtr Buffer, ref uint BufferSize);

    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerGetActiveScheme(IntPtr UserRootPowerKey, out IntPtr ActivePolicyGuid);

    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerSetActiveScheme(IntPtr UserRootPowerKey, ref Guid SchemeGuid);

    // --- Buttons/Lid + timeout interop (System Settings per-plan) ---
    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerReadACValueIndex(
        IntPtr RootPowerKey, ref Guid SchemeGuid, ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid, out uint AcValueIndex);

    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerReadDCValueIndex(
        IntPtr RootPowerKey, ref Guid SchemeGuid, ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid, out uint DcValueIndex);

    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerWriteACValueIndex(
        IntPtr RootPowerKey, ref Guid SchemeGuid, ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid, uint AcValueIndex);

    [DllImport("powrprof.dll", SetLastError = true)]
    private static extern uint PowerWriteDCValueIndex(
        IntPtr RootPowerKey, ref Guid SchemeGuid, ref Guid SubGroupOfPowerSettingsGuid,
        ref Guid PowerSettingGuid, uint DcValueIndex);

    [DllImport("kernel32.dll")]
    private static extern IntPtr LocalFree(IntPtr hMem);

    // Subgroup + setting GUIDs for the “System Settings” buttons/lid
    private static readonly Guid SUB_BUTTONS = new Guid("4f971e89-eebd-4455-a8de-9e59040e7347");
    private static readonly Guid SET_PBUTTON = new Guid("7648efa3-dd9c-4e3e-b566-50f929386280"); // Power button action
    private static readonly Guid SET_SBUTTON = new Guid("96996bc0-ad50-47ec-923b-6f41874dd9eb"); // Sleep button action
    private static readonly Guid SET_LID     = new Guid("5ca83367-6e45-459f-a27b-476b1d01c936"); // Lid close action

    // Subgroups for display & sleep
    private static readonly Guid SUB_VIDEO = new Guid("7516b95f-f776-4464-8c53-06167f40cc99");
    private static readonly Guid SUB_SLEEP = new Guid("238c9fa8-0aad-41ed-83f4-97be242c8f20");

    // Display timeout settings
    private static readonly Guid SET_VIDEOIDLE    = new Guid("3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"); // Turn off display after
    private static readonly Guid SET_VIDEOCONLOCK = new Guid("8ec4b3a5-6868-48c2-be75-4f3044be88a7"); // Console lock display off timeout

    // Sleep-related timeout settings
    private static readonly Guid SET_SLEEPIDLE     = new Guid("29f6c1db-86da-48c5-9fdb-f2b67b1f44da"); // Sleep after
    private static readonly Guid SET_HIBERNATEIDLE = new Guid("9d7815a6-7ee4-497e-8888-515a05f02364"); // Hibernate after
    private static readonly Guid SET_UNATTENDSLEEP = new Guid("7bc4a2f9-d8fc-4469-b07b-33eb785aaca0"); // Unattended sleep timeout

    // Allowed values (match Control Panel dropdowns)
    // 0=Do nothing, 1=Sleep, 2=Hibernate, 3=Shut down
    private enum ButtonLidAction : uint { DoNothing = 0, Sleep = 1, Hibernate = 2, Shutdown = 3 }

    // UI language
    private enum UiLanguage { English, Spanish }
    private UiLanguage uiLanguage = UiLanguage.English;
    private bool languageLoadedFromConfig = false;

    // Simple helper for bilingual strings
    private string L(string en, string es)
    {
        return (uiLanguage == UiLanguage.Spanish ? es : en);
    }

    // Embedded resource logical names (must match /resource:,"LogicalName")
    private const string RES_DESKTOP_DARK  = "Icon.Desktop.Dark.ico";
    private const string RES_DESKTOP_LIGHT = "Icon.Desktop.Light.ico";
    private const string RES_LAPTOP_DARK   = "Icon.Laptop.Dark.ico";
    private const string RES_LAPTOP_LIGHT  = "Icon.Laptop.Light.ico";
    private const string RES_BOLT_DARK     = "Icon.Bolt.Dark.ico";
    private const string RES_BOLT_LIGHT    = "Icon.Bolt.Light.ico";
    private const string RES_MOON_DARK     = "Icon.Moon.Dark.ico";
    private const string RES_MOON_LIGHT    = "Icon.Moon.Light.ico";

    internal const string AppId = "SwitchPowerTray";
    internal const string ShortcutName = "Switch Power Plan Tray.lnk";

    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SwitchPowerTray");
    private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.txt");
    private static readonly string LogPath    = Path.Combine(Path.GetTempPath(), "SwitchPowerTray.log");

    private readonly NotifyIcon tray;
    private readonly bool debug;

    private static volatile bool _blockLaunch = false;

    private string guidA = ""; // Desktop
    private string guidB = ""; // Laptop
    private string guidC = ""; // Bolt
    private string guidD = ""; // Moon

    private enum IconSet { Auto, Light, Dark }
    private IconSet iconSetPref = IconSet.Auto;

    private enum Slot { A_Desktop, B_Laptop, C_Bolt, D_Moon }

    private Icon exeIcon, lastIcon;
    private readonly Dictionary<string, Icon> icons = new Dictionary<string, Icon>(StringComparer.OrdinalIgnoreCase);

    private string activeGuid = "";
    private List<Plan> plans = new List<Plan>();

    private sealed class Plan
    {
        public string Guid;
        public string Name;
        public bool IsActive;
        public override string ToString()
        {
            return Name + " (" + Guid + (IsActive ? ", Active" : "") + ")";
        }
    }

    private struct AssignTag { public Slot Slot; public string Guid; public AssignTag(Slot s, string g) { Slot = s; Guid = g; } }

    private readonly Timer themePollTimer = new Timer();
    private bool _busy;

    // Hidden window to catch WM_QUERYENDSESSION / WM_ENDSESSION early
    private sealed class EndSessionWatcher : NativeWindow, IDisposable
    {
        private const int WM_QUERYENDSESSION = 0x0011;
        private const int WM_ENDSESSION      = 0x0016;

        public EndSessionWatcher()
        {
            CreateParams cp = new CreateParams();
            this.CreateHandle(cp);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_QUERYENDSESSION || m.Msg == WM_ENDSESSION)
            {
                TrayContext.BeginShutdown("WM_ENDSESSION");
            }
            base.WndProc(ref m);
        }

        public void Dispose()
        {
            try { if (this.Handle != IntPtr.Zero) this.DestroyHandle(); } catch { }
        }
    }
    private EndSessionWatcher endWatcher;

    public TrayContext(bool debugMode)
    {
        debug = debugMode;

        try { exeIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
        catch { exeIcon = SystemIcons.Application; }

        LoadAllIcons();
        LoadConfig();
        DetectLanguageIfNotSet();

        tray = new NotifyIcon();
        tray.Visible = false;
        tray.Text = L("Switch Power Plan", "Cambiar plan de energía");
        tray.ContextMenuStrip = BuildMenu();
        tray.MouseClick += OnTrayMouseClick;

        for (int i = 0; i < 5; i++)
        {
            RefreshPlansAndIcon();
            if (!string.IsNullOrEmpty(activeGuid)) break;
            Thread.Sleep(250);
        }
        tray.Visible = true;

        endWatcher = new EndSessionWatcher();

        SystemEvents.SessionEnding += (s, e) => { BeginShutdown("Context.SessionEnding"); };
        SystemEvents.SessionEnded  += (s, e) => { BeginShutdown("Context.SessionEnded"); };
        Application.ApplicationExit += (s, e) => { BeginShutdown("Context.ApplicationExit"); };

        themePollTimer.Interval = 2000;
        themePollTimer.Tick += OnThemePollTick;
        themePollTimer.Start();
    }

    private void DetectLanguageIfNotSet()
    {
        if (languageLoadedFromConfig) return;
        try
        {
            string ui = CultureInfo.InstalledUICulture.TwoLetterISOLanguageName;
            if (string.Equals(ui, "es", StringComparison.OrdinalIgnoreCase))
                uiLanguage = UiLanguage.Spanish;
            else
                uiLanguage = UiLanguage.English;
        }
        catch
        {
            uiLanguage = UiLanguage.English;
        }
    }

    internal static void BeginShutdown(string src)
    {
        _blockLaunch = true;
        try
        {
            File.AppendAllText(Path.Combine(Path.GetTempPath(), "SwitchPowerTray.log"),
                DateTime.Now.ToString("s") + "  BeginShutdown: " + src + Environment.NewLine);
        }
        catch { }
    }

    // ---------- Event handlers ----------
    private void OnTrayMouseClick(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left) ToggleToNextAssigned();
    }

    private void OnThemePollTick(object sender, EventArgs e)
    {
        if (_blockLaunch) { try { themePollTimer.Stop(); } catch { } return; }
        if (iconSetPref == IconSet.Auto) UpdateTrayIcon();
    }

    private void OnAssignClick(object sender, EventArgs e)
    {
        ToolStripMenuItem mi = sender as ToolStripMenuItem;
        if (mi == null || mi.Tag == null) return;
        AssignTag t = (AssignTag)mi.Tag;
        SetSlot(t.Slot, t.Guid);
        SaveConfig();
        UpdateTrayIcon();
    }

    private void OnAssignClearClick(object sender, EventArgs e)
    {
        ToolStripMenuItem mi = sender as ToolStripMenuItem;
        if (mi == null || mi.Tag == null) return;
        Slot slot = (Slot)mi.Tag;
        SetSlot(slot, "");
        SaveConfig();
        UpdateTrayIcon();
    }

    private void OnSwitchToClick(object sender, EventArgs e)
    {
        ToolStripMenuItem mi = sender as ToolStripMenuItem;
        if (mi == null || mi.Tag == null) return;
        string guid = mi.Tag as string;
        if (!string.IsNullOrEmpty(guid)) TrySetActive(guid);
    }

    private void OnThemeAuto(object sender, EventArgs e)  { iconSetPref = IconSet.Auto;  SaveConfig(); UpdateTrayIcon(); }
    private void OnThemeLight(object sender, EventArgs e) { iconSetPref = IconSet.Light; SaveConfig(); UpdateTrayIcon(); }
    private void OnThemeDark (object sender, EventArgs e) { iconSetPref = IconSet.Dark;  SaveConfig(); UpdateTrayIcon(); }

    private void OnLanguageEnglish(object sender, EventArgs e)
    {
        uiLanguage = UiLanguage.English;
        SaveConfig();
        RebuildMenu();
    }

    private void OnLanguageSpanish(object sender, EventArgs e)
    {
        uiLanguage = UiLanguage.Spanish;
        SaveConfig();
        RebuildMenu();
    }

    private void RebuildMenu()
    {
        if (tray == null) return;
        tray.ContextMenuStrip = BuildMenu();
        tray.Text = L("Switch Power Plan", "Cambiar plan de energía");
    }

    private void OnOpenPowerOptions(object sender, EventArgs e)
    {
        try { Process.Start("control.exe", "powercfg.cpl"); } catch { }
    }

    private void OnExit(object sender, EventArgs e)
    {
        ExitThread();
    }

    // Helper: don’t close dropdown when mouse is still inside it
    private void WireNoCloseOnItemClick(ToolStripDropDown drop)
    {
        if (drop == null) return;

        drop.Closing += (s, e) =>
        {
            // If mouse is inside this dropdown’s bounds, keep it open.
            Point p = Cursor.Position; // screen coords
            if (drop.Bounds.Contains(p))
            {
                e.Cancel = true;
            }
        };
    }

    // ---------- Menu construction ----------
    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();

        menu.Items.Add(new ToolStripMenuItem(
            L("Toggle now", "Cambiar ahora"), null, (EventHandler)OnToggleNow));

        menu.Items.Add(BuildAssignSubmenu(
            L("Assign Slot A (Desktop icon) →", "Asignar ranura A (Ícono-Escritorio) →"), Slot.A_Desktop));
        menu.Items.Add(BuildAssignSubmenu(
            L("Assign Slot B (Laptop icon) →", "Asignar ranura B (Ícono-Portátil) →"),  Slot.B_Laptop));
        menu.Items.Add(BuildAssignSubmenu(
            L("Assign Slot C (Bolt icon) →",   "Asignar ranura C (Ícono-Rayo) →"),       Slot.C_Bolt));
        menu.Items.Add(BuildAssignSubmenu(
            L("Assign Slot D (Moon icon) →",   "Asignar ranura D (Ícono-Luna) →"),       Slot.D_Moon));

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(BuildSwitchToSubmenu());

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(BuildThemeMenu());

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(BuildLanguageMenu());

        // per-plan “System Settings” editor (power button / sleep button / lid)
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(BuildCustomizeButtonsMenu());

        // per-plan Display & Sleep timeouts
        menu.Items.Add(BuildCustomizeDisplaySleepMenu());

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(new ToolStripMenuItem(
            L("Open Power Options…", "Abrir opciones de energía…"), null, OnOpenPowerOptions));

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(new ToolStripMenuItem(L("Exit (Close Program)", "Salir (Cerrar Programa)"), null, OnExit));

        // root dropdown
        WireNoCloseOnItemClick(menu);

        return menu;
    }

    private ToolStripMenuItem BuildLanguageMenu()
    {
        var sub = new ToolStripMenuItem(L("Language", "Idioma"));
        sub.DropDownOpening += delegate
        {
            sub.DropDownItems.Clear();
            var enItem = new ToolStripMenuItem("English", null, OnLanguageEnglish);
            var esItem = new ToolStripMenuItem("Español", null, OnLanguageSpanish);

            enItem.Checked = (uiLanguage == UiLanguage.English);
            esItem.Checked = (uiLanguage == UiLanguage.Spanish);

            sub.DropDownItems.Add(enItem);
            sub.DropDownItems.Add(esItem);
        };

        WireNoCloseOnItemClick(sub.DropDown);
        return sub;
    }

    private void OnToggleNow(object sender, EventArgs e)
    {
        ToggleToNextAssigned();
    }

    private ToolStripMenuItem BuildAssignSubmenu(string title, Slot slot)
    {
        var sub = new ToolStripMenuItem(title);
        sub.DropDownOpening += delegate
        {
            sub.DropDownItems.Clear();
            EnsurePlanList();

            foreach (Plan p in plans)
            {
                string label = p.Name + (p.IsActive ? "  (" + L("Active", "Activo") + ")" : "");
                var item = new ToolStripMenuItem(label, null, OnAssignClick);
                item.Tag = new AssignTag(slot, p.Guid);

                if (string.Equals(GetSlot(slot), p.Guid, StringComparison.OrdinalIgnoreCase))
                    item.Checked = true;

                sub.DropDownItems.Add(item);
            }

            if (plans.Count > 0) sub.DropDownItems.Add(new ToolStripSeparator());
            var clear = new ToolStripMenuItem(L("Clear this slot", "Limpiar esta ranura"), null, OnAssignClearClick);
            clear.Tag = slot;
            sub.DropDownItems.Add(clear);
        };

        WireNoCloseOnItemClick(sub.DropDown);
        return sub;
    }

    private ToolStripMenuItem BuildThemeMenu()
    {
        var sub = new ToolStripMenuItem(L("Icon contrast", "Contraste de iconos"));
        sub.DropDownOpening += delegate
        {
            sub.DropDownItems.Clear();
            var iAuto  = new ToolStripMenuItem(
                L("Auto (match system, high contrast)", "Auto (según sistema, alto contraste)"),
                null, OnThemeAuto);
            var iLight = new ToolStripMenuItem(
                L("Use Light icons", "Usar iconos claros"),
                null, OnThemeLight);
            var iDark  = new ToolStripMenuItem(
                L("Use Dark icons", "Usar iconos oscuros"),
                null, OnThemeDark);

            iAuto.Checked  = (iconSetPref == IconSet.Auto);
            iLight.Checked = (iconSetPref == IconSet.Light);
            iDark.Checked  = (iconSetPref == IconSet.Dark);

            sub.DropDownItems.Add(iAuto);
            sub.DropDownItems.Add(iLight);
            sub.DropDownItems.Add(iDark);
        };

        WireNoCloseOnItemClick(sub.DropDown);
        return sub;
    }

    private ToolStripMenuItem BuildSwitchToSubmenu()
    {
        var sub = new ToolStripMenuItem(L("Switch to…", "Cambiar a…"));
        sub.DropDownOpening += delegate
        {
            sub.DropDownItems.Clear();
            EnsurePlanList();
            AddSwitchItem(sub, "A (" + L("Desktop", "Escritorio") + ")", guidA);
            AddSwitchItem(sub, "B (" + L("Laptop", "Portátil")   + ")",  guidB);
            AddSwitchItem(sub, "C (" + L("Bolt", "Rayo")         + ")",  guidC);
            AddSwitchItem(sub, "D (" + L("Moon", "Luna")         + ")",  guidD);
            if (sub.DropDownItems.Count == 0)
                sub.DropDownItems.Add(L("(no slots assigned)", "(sin ranuras asignadas)"));
        };

        WireNoCloseOnItemClick(sub.DropDown);
        return sub;
    }

    private void AddSwitchItem(ToolStripMenuItem parent, string caption, string guid)
    {
        if (string.IsNullOrEmpty(guid)) return;
        string name = FindPlanName(guid);
        var item = new ToolStripMenuItem(caption + ": " + name, null, OnSwitchToClick);
        item.Tag = guid;
        if (!string.IsNullOrEmpty(activeGuid) &&
            string.Equals(activeGuid, guid, StringComparison.OrdinalIgnoreCase)) item.Checked = true;
        parent.DropDownItems.Add(item);
    }

    // ---------- Buttons & Lid customizer ----------
    private ToolStripMenuItem BuildCustomizeButtonsMenu()
    {
        var root = new ToolStripMenuItem(L("Customize (Buttons && Lid) →", "Personalizar (Botones y tapa) →"));

        root.DropDownOpening += delegate
        {
            root.DropDownItems.Clear();
            EnsurePlanList();

            foreach (var p in plans)
            {
                var planItem = new ToolStripMenuItem(
                    p.Name + (p.IsActive ? "  (" + L("Active", "Activo") + ")" : ""));
                planItem.Tag = p.Guid;

                Guid scheme = Guid.Parse(p.Guid);
                planItem.DropDownItems.Add(
                    BuildSettingMenu(L("Power button", "Botón de encendido"), scheme, SET_PBUTTON));
                planItem.DropDownItems.Add(
                    BuildSettingMenu(L("Sleep button", "Botón de suspensión"), scheme, SET_SBUTTON));
                planItem.DropDownItems.Add(
                    BuildSettingMenu(L("Closing lid", "Cerrar tapa"),         scheme, SET_LID));

                WireNoCloseOnItemClick(planItem.DropDown);
                root.DropDownItems.Add(planItem);
            }

            if (root.DropDownItems.Count == 0)
                root.DropDownItems.Add(L("(no power plans found)", "(no se encontraron planes de energía)"));
        };

        WireNoCloseOnItemClick(root.DropDown);
        return root;
    }

    private ToolStripMenuItem BuildSettingMenu(string caption, Guid scheme, Guid setting)
    {
        var settingItem = new ToolStripMenuItem(caption + " →");

        settingItem.DropDownOpening += delegate
        {
            settingItem.DropDownItems.Clear();

            // AC submenu
            var acMenu = new ToolStripMenuItem(L("On AC →", "Con corriente →"));
            AddActionChoiceItems(acMenu, scheme, setting, true);
            WireNoCloseOnItemClick(acMenu.DropDown);
            settingItem.DropDownItems.Add(acMenu);

            // DC submenu
            var dcMenu = new ToolStripMenuItem(L("On battery →", "Con batería →"));
            AddActionChoiceItems(dcMenu, scheme, setting, false);
            WireNoCloseOnItemClick(dcMenu.DropDown);
            settingItem.DropDownItems.Add(dcMenu);
        };

        WireNoCloseOnItemClick(settingItem.DropDown);
        return settingItem;
    }

    private void AddActionChoiceItems(ToolStripMenuItem parent, Guid scheme, Guid setting, bool ac)
    {
        parent.DropDownItems.Clear();

        uint current = ReadAction(scheme, setting, ac);

        Action<uint> add = delegate (uint val)
        {
            var mi = new ToolStripMenuItem(ActionName(val));
            if (current == val) mi.Checked = true;

            mi.Click += delegate
            {
                // apply change
                WriteAction(scheme, setting, ac, val);

                // update checkmarks in this submenu
                foreach (ToolStripItem tsi in parent.DropDownItems)
                {
                    ToolStripMenuItem tmi = tsi as ToolStripMenuItem;
                    if (tmi != null)
                        tmi.Checked = (tmi == mi);
                }
            };

            parent.DropDownItems.Add(mi);
        };

        add((uint)ButtonLidAction.DoNothing);
        add((uint)ButtonLidAction.Sleep);
        add((uint)ButtonLidAction.Hibernate);
        add((uint)ButtonLidAction.Shutdown);
    }

    private static uint ReadAction(Guid scheme, Guid setting, bool ac)
    {
        try
        {
            uint v;
            Guid sub = SUB_BUTTONS;
            Guid set = setting;
            if (ac)
            {
                if (PowerReadACValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, out v) == 0) return v;
            }
            else
            {
                if (PowerReadDCValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, out v) == 0) return v;
            }
        }
        catch { }
        return 0; // default
    }

    private void WriteAction(Guid scheme, Guid setting, bool ac, uint value)
    {
        try
        {
            Guid sub = SUB_BUTTONS;
            Guid set = setting;
            if (ac) PowerWriteACValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, value);
            else    PowerWriteDCValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, value);

            // If we edited the active scheme, re-apply it so changes take effect immediately
            if (!string.IsNullOrEmpty(activeGuid) &&
                string.Equals(scheme.ToString(), activeGuid, StringComparison.OrdinalIgnoreCase))
            {
                var g = scheme;
                try { PowerSetActiveScheme(IntPtr.Zero, ref g); } catch { }
            }
        }
        catch { }
    }

    private string ActionName(uint v)
    {
        switch ((ButtonLidAction)v)
        {
            case ButtonLidAction.DoNothing: return L("Do nothing", "No hacer nada");
            case ButtonLidAction.Sleep:     return L("Sleep", "Suspender");
            case ButtonLidAction.Hibernate: return L("Hibernate", "Hibernar");
            case ButtonLidAction.Shutdown:  return L("Shut down", "Apagar");
            default: return v.ToString();
        }
    }

    // ---------- Display & Sleep customizer ----------
    private ToolStripMenuItem BuildCustomizeDisplaySleepMenu()
    {
        var root = new ToolStripMenuItem(L(
            "Customize (Display && Sleep) →",
            "Personalizar (Pantalla y suspensión) →"));

        root.DropDownOpening += delegate
        {
            root.DropDownItems.Clear();
            EnsurePlanList();

            foreach (var p in plans)
            {
                var planItem = new ToolStripMenuItem(
                    p.Name + (p.IsActive ? "  (" + L("Active", "Activo") + ")" : ""));
                planItem.Tag = p.Guid;

                Guid scheme = Guid.Parse(p.Guid);

                // Display group
                planItem.DropDownItems.Add(
                    BuildTimeoutMenu(
                        L("Display off timeout", "Tiempo para apagar pantalla"),
                        scheme, SUB_VIDEO, SET_VIDEOIDLE));
                planItem.DropDownItems.Add(
                    BuildTimeoutMenu(
                        L("Console lock display off timeout", "Tiempo de pantalla al bloquear"),
                        scheme, SUB_VIDEO, SET_VIDEOCONLOCK));
                planItem.DropDownItems.Add(new ToolStripSeparator());

                // Sleep group
                planItem.DropDownItems.Add(
                    BuildTimeoutMenu(
                        L("Sleep after", "Suspender después de"),
                        scheme, SUB_SLEEP, SET_SLEEPIDLE));
                planItem.DropDownItems.Add(
                    BuildTimeoutMenu(
                        L("Hibernate after", "Hibernar después de"),
                        scheme, SUB_SLEEP, SET_HIBERNATEIDLE));
                planItem.DropDownItems.Add(
                    BuildTimeoutMenu(
                        L("Unattended sleep timeout", "Tiempo de suspensión desatendida"),
                        scheme, SUB_SLEEP, SET_UNATTENDSLEEP));

                WireNoCloseOnItemClick(planItem.DropDown);
                root.DropDownItems.Add(planItem);
            }

            if (root.DropDownItems.Count == 0)
                root.DropDownItems.Add(L("(no power plans found)", "(no se encontraron planes de energía)"));
        };

        WireNoCloseOnItemClick(root.DropDown);
        return root;
    }

    private ToolStripMenuItem BuildTimeoutMenu(string caption, Guid scheme, Guid subgroup, Guid setting)
    {
        var settingItem = new ToolStripMenuItem(caption + " →");

        settingItem.DropDownOpening += delegate
        {
            settingItem.DropDownItems.Clear();

            // AC submenu
            var acMenu = new ToolStripMenuItem(L("On AC →", "Con corriente →"));
            AddTimeoutChoiceItems(acMenu, scheme, subgroup, setting, true);
            WireNoCloseOnItemClick(acMenu.DropDown);
            settingItem.DropDownItems.Add(acMenu);

            // DC submenu
            var dcMenu = new ToolStripMenuItem(L("On battery →", "Con batería →"));
            AddTimeoutChoiceItems(dcMenu, scheme, subgroup, setting, false);
            WireNoCloseOnItemClick(dcMenu.DropDown);
            settingItem.DropDownItems.Add(dcMenu);
        };

        WireNoCloseOnItemClick(settingItem.DropDown);
        return settingItem;
    }

    private void AddTimeoutChoiceItems(
        ToolStripMenuItem parent,
        Guid scheme,
        Guid subgroup,
        Guid setting,
        bool ac)
    {
        parent.DropDownItems.Clear();

        uint current = ReadTimeoutSeconds(scheme, subgroup, setting, ac);

        // minutes list: 0 = Never, then some common values
        int[] mins = new[] { 0, 1, 2, 3, 5, 10, 15, 20, 30, 60, 120 };

        foreach (int m in mins)
        {
            uint secs = (uint)(m * 60);
            string label;
            if (m == 0)
            {
                label = L("Never", "Nunca");
            }
            else if (m < 60)
            {
                label = m + " " + L("min", "min");
            }
            else
            {
                int hrs = m / 60;
                if (hrs == 1)
                    label = "1 " + L("hour", "hora");
                else
                    label = hrs + " " + L("hours", "horas");
            }

            uint secsLocal = secs; // capture for closure

            var mi = new ToolStripMenuItem(label);
            mi.Tag = secsLocal; // store raw seconds for later comparison

            if (current == secsLocal) mi.Checked = true;

            mi.Click += delegate
            {
                WriteTimeoutSeconds(scheme, subgroup, setting, ac, secsLocal);
                current = secsLocal;

                // refresh checkmarks in this submenu
                foreach (ToolStripItem tsi in parent.DropDownItems)
                {
                    ToolStripMenuItem tmi = tsi as ToolStripMenuItem;
                    if (tmi != null && tmi.Tag != null && tmi.Tag is uint)
                    {
                        uint v = (uint)tmi.Tag;
                        tmi.Checked = (v == secsLocal);
                    }
                }
            };

            parent.DropDownItems.Add(mi);
        }

        // Separator + Custom…
        parent.DropDownItems.Add(new ToolStripSeparator());

        var customItem = new ToolStripMenuItem(L("Custom…", "Personalizado…"));
        customItem.Click += delegate
        {
            uint newSeconds = PromptCustomTimeoutSeconds(current);
            if (newSeconds == current) return; // user cancelled or no change

            WriteTimeoutSeconds(scheme, subgroup, setting, ac, newSeconds);
            current = newSeconds;

            // update checkmarks: match only exact preset if exists
            foreach (ToolStripItem tsi in parent.DropDownItems)
            {
                ToolStripMenuItem tmi = tsi as ToolStripMenuItem;
                if (tmi != null && tmi.Tag != null && tmi.Tag is uint)
                {
                    uint v = (uint)tmi.Tag;
                    tmi.Checked = (v == newSeconds);
                }
            }
        };

        parent.DropDownItems.Add(customItem);
    }

    private static uint ReadTimeoutSeconds(Guid scheme, Guid subgroup, Guid setting, bool ac)
    {
        try
        {
            uint v;
            Guid sub = subgroup;
            Guid set = setting;

            if (ac)
            {
                if (PowerReadACValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, out v) == 0)
                    return v;
            }
            else
            {
                if (PowerReadDCValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, out v) == 0)
                    return v;
            }
        }
        catch { }

        // default / not found
        return 0;
    }

    private void WriteTimeoutSeconds(Guid scheme, Guid subgroup, Guid setting, bool ac, uint seconds)
    {
        try
        {
            Guid sub = subgroup;
            Guid set = setting;

            if (ac)
                PowerWriteACValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, seconds);
            else
                PowerWriteDCValueIndex(IntPtr.Zero, ref scheme, ref sub, ref set, seconds);

            // If we edited the active scheme, re-apply it so changes take effect immediately
            if (!string.IsNullOrEmpty(activeGuid) &&
                string.Equals(scheme.ToString(), activeGuid, StringComparison.OrdinalIgnoreCase))
            {
                var g = scheme;
                try { PowerSetActiveScheme(IntPtr.Zero, ref g); } catch { }
            }
        }
        catch { }
    }

    // Simple custom-minutes dialog (returns seconds)
    private uint PromptCustomTimeoutSeconds(uint currentSeconds)
    {
        uint currentMinutes = currentSeconds / 60;
        uint resultMinutes = currentMinutes;

        using (Form f = new Form())
        {
            f.Text = L("Custom timeout", "Tiempo personalizado");
            f.FormBorderStyle = FormBorderStyle.FixedDialog;
            f.StartPosition = FormStartPosition.CenterScreen;
            f.MinimizeBox = false;
            f.MaximizeBox = false;
            f.ClientSize = new Size(320, 130);

            Label lbl = new Label
            {
                Left = 9,
                Top = 9,
                Width = 300,
                Text = L("Enter minutes (0 = Never):", "Ingresa minutos (0 = Nunca):")
            };

            TextBox tb = new TextBox
            {
                Left = 12,
                Top = 35,
                Width = 100,
                Text = currentMinutes.ToString()
            };

            Button ok = new Button
            {
                Text = "OK",
                Left = 90,
                Top = 75,
                Width = 80,
                DialogResult = DialogResult.None
            };

            Button cancel = new Button
            {
                Text = L("Cancel", "Cancelar"),
                Left = 180,
                Top = 75,
                Width = 80,
                DialogResult = DialogResult.Cancel
            };

            ok.Click += delegate
            {
                string text = tb.Text.Trim();
                if (string.IsNullOrEmpty(text))
                {
                    MessageBox.Show(
                        f,
                        L("Please enter a non-negative integer number of minutes.", "Ingresa un número entero de minutos mayor o igual a 0."),
                        L("Invalid value", "Valor inválido"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    tb.Focus();
                    tb.SelectAll();
                    return;
                }

                uint m;
                if (!uint.TryParse(text, out m))
                {
                    MessageBox.Show(
                        f,
                        L("Please enter a non-negative integer number of minutes.", "Ingresa un número entero de minutos mayor o igual a 0."),
                        L("Invalid value", "Valor inválido"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    tb.Focus();
                    tb.SelectAll();
                    return;
                }

                resultMinutes = m;
                f.DialogResult = DialogResult.OK;
                f.Close();
            };

            f.Controls.Add(lbl);
            f.Controls.Add(tb);
            f.Controls.Add(ok);
            f.Controls.Add(cancel);

            f.AcceptButton = ok;
            f.CancelButton = cancel;

            if (f.ShowDialog() == DialogResult.OK)
            {
                return resultMinutes * 60;
            }
            else
            {
                // Cancel → keep current
                return currentSeconds;
            }
        }
    }

    // ---------- Core logic ----------
    private void ToggleToNextAssigned()
    {
        if (_busy) return;
        _busy = true;
        try
        {
            EnsurePlanList();

            List<string> cycle = new List<string>();
            if (!string.IsNullOrEmpty(guidA)) cycle.Add(guidA);
            if (!string.IsNullOrEmpty(guidB)) cycle.Add(guidB);
            if (!string.IsNullOrEmpty(guidC)) cycle.Add(guidC);
            if (!string.IsNullOrEmpty(guidD)) cycle.Add(guidD);

            if (cycle.Count == 0)
            {
                MessageBox.Show(
                    L("Assign at least one slot in the tray menu first.",
                      "Asigna al menos una ranura en el menú de la bandeja primero."),
                    L("Switch Power Plan", "Cambiar plan de energía"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string target;
            if (string.IsNullOrEmpty(activeGuid))
            {
                target = cycle[0];
            }
            else
            {
                int idx = -1;
                for (int i = 0; i < cycle.Count; i++)
                    if (string.Equals(cycle[i], activeGuid, StringComparison.OrdinalIgnoreCase)) { idx = i; break; }
                target = (idx >= 0 && idx + 1 < cycle.Count) ? cycle[idx + 1] : cycle[0];
            }

            TrySetActive(target);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message,
                L("Toggle error", "Error al cambiar"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally { _busy = false; }
    }

    private void TrySetActive(string guid)
    {
        if (string.IsNullOrEmpty(guid)) return;

        // show the destination icon immediately
        PreselectIconForGuid(guid);

        // Bail if OS is shutting down / logging off / TS session
        if (!_blockLaunch && !Environment.HasShutdownStarted && !SystemInformation.TerminalServerSession)
        {
            Guid g;
            if (Guid.TryParse(guid, out g))
            {
                try { PowerSetActiveScheme(IntPtr.Zero, ref g); } catch { }
            }
        }

        // resync name + icon
        RefreshPlansAndIcon();
    }

    private void PreselectIconForGuid(string guid)
    {
        Icon icon = IconForGuid(guid);
        if (icon != null)
        {
            tray.Icon = icon;
            lastIcon = icon;
        }
    }

    private void RefreshPlansAndIcon()
    {
        EnsurePlanList();
        UpdateTrayIcon();
    }

    private void EnsurePlanList()
    {
        plans = ListPlans();
        activeGuid = GetActiveSchemeGuid();
    }

    private void UpdateTrayIcon()
    {
        Icon icon = exeIcon;

        if (!string.IsNullOrEmpty(activeGuid))
        {
            Icon chosen = IconForGuid(activeGuid);
            if (chosen != null) icon = chosen;
        }

        if (icon == null) icon = (lastIcon != null ? lastIcon : (exeIcon != null ? exeIcon : SystemIcons.Application));

        tray.Icon = icon;
        lastIcon = icon;

        string activeName = FindPlanName(activeGuid ?? "");
        tray.Text = string.IsNullOrEmpty(activeName)
            ? L("Switch Power Plan", "Cambiar plan de energía")
            : TrimForTray(activeName);
    }

    private Icon IconForGuid(string guid)
    {
        string slotName = null;
        if (!string.IsNullOrEmpty(guidA) && guid.Equals(guidA, StringComparison.OrdinalIgnoreCase)) slotName = "Desktop";
        else if (!string.IsNullOrEmpty(guidB) && guid.Equals(guidB, StringComparison.OrdinalIgnoreCase)) slotName = "Laptop";
        else if (!string.IsNullOrEmpty(guidC) && guid.Equals(guidC, StringComparison.OrdinalIgnoreCase)) slotName = "Bolt";
        else if (!string.IsNullOrEmpty(guidD) && guid.Equals(guidD, StringComparison.OrdinalIgnoreCase)) slotName = "Moon";
        if (slotName == null) return null;

        string variant;
        if (iconSetPref == IconSet.Auto)
            variant = SystemIsLight() ? "Dark" : "Light"; // Light system => use Dark icons (contrast)
        else
            variant = (iconSetPref == IconSet.Light) ? "Light" : "Dark";

        Icon ic;
        if (icons.TryGetValue(slotName + "." + variant, out ic))
            return ic;
        return null;
    }

    private static string TrimForTray(string s)
    {
        if (s == null) return "";
        s = s.Replace("\r", "").Replace("\n", " · ");
        if (s.Length > 63) s = s.Substring(0, 63);
        return s;
    }

    private string FindPlanName(string guid)
    {
        if (string.IsNullOrEmpty(guid)) return "";
        EnsurePlanList(); // keeps tooltip fresh
        foreach (Plan p in plans)
            if (string.Equals(p.Guid, guid, StringComparison.OrdinalIgnoreCase))
                return p.Name;
        return guid;
    }

    // === PowrProf-based plan listing ===
    private static List<Plan> ListPlans()
    {
        var list = new List<Plan>();
        if (_blockLaunch || Environment.HasShutdownStarted) return list;

        string active = GetActiveSchemeGuid();

        uint index = 0;
        uint size = (uint)Marshal.SizeOf(typeof(Guid));
        IntPtr buf = Marshal.AllocHGlobal((int)size);
        try
        {
            while (true)
            {
                uint s = size;
                uint res = PowerEnumerate(IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, ACCESS_SCHEME, index, buf, ref s);
                if (res == ERROR_NO_MORE_ITEMS) break;
                if (res != 0) break;

                Guid g = (Guid)Marshal.PtrToStructure(buf, typeof(Guid));
                string name = ReadFriendlyName(g);
                var p = new Plan
                {
                    Guid = g.ToString(),
                    Name = !string.IsNullOrEmpty(name) ? name : g.ToString(),
                    IsActive = string.Equals(active, g.ToString(), StringComparison.OrdinalIgnoreCase)
                };
                list.Add(p);

                index++;
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buf);
        }
        return list;
    }

    private static string ReadFriendlyName(Guid scheme)
    {
        uint needed = 0;
        uint res = PowerReadFriendlyName(IntPtr.Zero, ref scheme, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, ref needed);
        if (res != 0 || needed == 0) return null;

        IntPtr mem = Marshal.AllocHGlobal((int)needed);
        try
        {
            res = PowerReadFriendlyName(IntPtr.Zero, ref scheme, IntPtr.Zero, IntPtr.Zero, mem, ref needed);
            if (res != 0) return null;
            return Marshal.PtrToStringUni(mem);
        }
        finally
        {
            Marshal.FreeHGlobal(mem);
        }
    }

    private static string GetActiveSchemeGuid()
    {
        try
        {
            IntPtr ptr;
            uint res = PowerGetActiveScheme(IntPtr.Zero, out ptr);
            if (res != 0 || ptr == IntPtr.Zero) return "";
            Guid g = (Guid)Marshal.PtrToStructure(ptr, typeof(Guid));
            LocalFree(ptr);
            return g.ToString();
        }
        catch { return ""; }
    }

    internal static void LogAndShow(string where, Exception ex)
    {
        try
        {
            File.AppendAllText(LogPath, string.Format(
                "==== {0} {1} ====\r\n{2}: {3}\r\n{4}\r\n\r\n",
                DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString(),
                where, ex != null ? ex.Message : "(null)",
                ex != null ? ex.ToString() : "(no stack)"));
            MessageBox.Show((ex != null ? ex.Message : "(no exception)") + "\r\n\r\nLog: " + LogPath,
                "SwitchPowerTray error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch { }
    }

    protected override void ExitThreadCore()
    {
        BeginShutdown("Context.ExitThreadCore");

        try
        {
            if (endWatcher != null) { endWatcher.Dispose(); endWatcher = null; }
        }
        catch { }

        try
        {
            themePollTimer.Stop();
            tray.Visible = false;
            tray.Dispose();
            if (exeIcon != null) exeIcon.Dispose();
            foreach (KeyValuePair<string, Icon> kv in icons) if (kv.Value != null) kv.Value.Dispose();
        }
        catch { }
        base.ExitThreadCore();
    }

    private static void LogTrace(string msg)
    {
        try
        {
            File.AppendAllText(LogPath, DateTime.Now.ToString("s") + "  " + msg + Environment.NewLine);
        }
        catch { }
    }

    private void LoadAllIcons()
    {
        AddIcon("Desktop.Dark", RES_DESKTOP_DARK);
        AddIcon("Desktop.Light", RES_DESKTOP_LIGHT);
        AddIcon("Laptop.Dark",  RES_LAPTOP_DARK);
        AddIcon("Laptop.Light", RES_LAPTOP_LIGHT);
        AddIcon("Bolt.Dark",    RES_BOLT_DARK);
        AddIcon("Bolt.Light",   RES_BOLT_LIGHT);
        AddIcon("Moon.Dark",    RES_MOON_DARK);
        AddIcon("Moon.Light",   RES_MOON_LIGHT);
    }

    private void AddIcon(string key, string resName)
    {
        Icon ic = LoadEmbeddedIcon(resName);
        if (ic != null) icons[key] = ic;
    }

    private static Icon LoadEmbeddedIcon(string logicalName)
    {
        try
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream s = asm.GetManifestResourceStream(logicalName))
            {
                if (s != null) return new Icon(s);
            }
        }
        catch { }
        return null;
    }

    private bool SystemIsLight()
    {
        try
        {
            using (RegistryKey p = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
            {
                if (p != null)
                {
                    object v = p.GetValue("AppsUseLightTheme");
                    if (v is int) return ((int)v) != 0; // 1=Light, 0=Dark
                }
            }
        }
        catch { }
        return true;
    }

    private void LoadConfig()
    {
        try
        {
            if (!File.Exists(ConfigPath)) return;
            foreach (string line in File.ReadAllLines(ConfigPath))
            {
                string[] kv = line.Split(new char[] { '=' }, 2);
                if (kv.Length != 2) continue;
                string k = kv[0].Trim().ToUpperInvariant();
                string v = kv[1].Trim();
                if (k == "GUIDA") guidA = v;
                else if (k == "GUIDB") guidB = v;
                else if (k == "GUIDC") guidC = v;
                else if (k == "GUIDD") guidD = v;
                else if (k == "ICONSET")
                {
                    try { iconSetPref = (IconSet)Enum.Parse(typeof(IconSet), v, true); }
                    catch { }
                }
                else if (k == "LANG")
                {
                    languageLoadedFromConfig = true;
                    if (string.Equals(v, "Spanish", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(v, "Español", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(v, "Es", StringComparison.OrdinalIgnoreCase))
                        uiLanguage = UiLanguage.Spanish;
                    else
                        uiLanguage = UiLanguage.English;
                }
            }
        }
        catch { }
    }

    private void SaveConfig()
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            File.WriteAllLines(ConfigPath, new string[]
            {
                "GUIDA=" + (guidA ?? ""),
                "GUIDB=" + (guidB ?? ""),
                "GUIDC=" + (guidC ?? ""),
                "GUIDD=" + (guidD ?? ""),
                "ICONSET=" + iconSetPref.ToString(),
                "LANG=" + uiLanguage.ToString()
            });
        }
        catch { }
    }

    private string GetSlot(Slot s)
    {
        if (s == Slot.A_Desktop) return guidA;
        if (s == Slot.B_Laptop)  return guidB;
        if (s == Slot.C_Bolt)    return guidC;
        if (s == Slot.D_Moon)    return guidD;
        return "";
    }

    private void SetSlot(Slot s, string guid)
    {
        if (s == Slot.A_Desktop)      guidA = guid;
        else if (s == Slot.B_Laptop)  guidB = guid;
        else if (s == Slot.C_Bolt)    guidC = guid;
        else if (s == Slot.D_Moon)    guidD = guid;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SetCurrentProcessExplicitAppUserModelID(string AppID);

    internal static void EnsureIdentityAndShortcut()
    {
        try { SetCurrentProcessExplicitAppUserModelID(AppId); } catch { }

        string programs = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
        string lnkPath = Path.Combine(programs, ShortcutName);
        string exePath = Application.ExecutablePath;
        try { CreateOrUpdateShortcut(lnkPath, exePath, exePath, AppId); } catch { }
    }

    // === Shell link COM interop + helper ===
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")]
    private interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int iIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    [ComImport, Guid("0000010B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPersistFile
    {
        void GetClassID(out Guid pClassID);
        void IsDirty();
        void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
        void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);
        void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
        void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
    }

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore
    {
        void GetCount(out uint cProps);
        void GetAt(uint iProp, out PROPERTYKEY pkey);
        void GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        void SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        void Commit();
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    private struct PROPERTYKEY { public Guid fmtid; public uint pid; }

    [StructLayout(LayoutKind.Explicit)]
    private struct PROPVARIANT
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pointerValue;
        public static PROPVARIANT FromString(string s)
        {
            PROPVARIANT v = new PROPVARIANT();
            v.vt = 31; // VT_LPWSTR
            v.pointerValue = Marshal.StringToCoTaskMemUni(s);
            return v;
        }
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    private class CShellLink { }

    private static readonly PROPERTYKEY PKEY_AppUserModel_ID = new PROPERTYKEY
    {
        fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"),
        pid = 5
    };

    private static void CreateOrUpdateShortcut(string lnkPath, string exePath, string iconPath, string appId)
    {
        IShellLinkW link = (IShellLinkW)new CShellLink();
        link.SetPath(exePath);
        link.SetDescription("Switch Power Plan Tray");
        link.SetIconLocation(iconPath, 0);

        IPropertyStore props = (IPropertyStore)link;
        PROPVARIANT pv = PROPVARIANT.FromString(appId);
        PROPERTYKEY key = PKEY_AppUserModel_ID;
        props.SetValue(ref key, ref pv);
        props.Commit();

        IPersistFile file = (IPersistFile)link;
        file.Save(lnkPath, true);
    }
}
