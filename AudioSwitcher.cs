using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AudioSwitcherApp
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var app = new AudioSwitcherApplication();
            app.Run();
        }
    }

    internal sealed class AudioSwitcherApplication
    {
        private readonly string configPath;
        private readonly NotifyIcon notifyIcon;
        private readonly HotkeyWindow window;
        private readonly AppIconResources iconResources;
        private AppConfig config;
        private List<AudioDeviceInfo> configuredDevices;
        private int currentDeviceIndex;

        public AudioSwitcherApplication()
        {
            var exeDir = Path.GetDirectoryName(Application.ExecutablePath) ?? AppDomain.CurrentDomain.BaseDirectory;
            configPath = Path.Combine(exeDir, "config.json");

            iconResources = AppIconResources.Create(exeDir);
            notifyIcon = new NotifyIcon
            {
                Icon = iconResources.Icon,
                Text = "Audio Switcher",
                Visible = true
            };

            window = new HotkeyWindow();
            window.HotkeyPressed += delegate { SwitchToNextDevice(); };
        }

        public void Run()
        {
            try
            {
                config = LoadOrCreateConfig(configPath);
                ConfigureTrayMenu();

                if (!ApplyConfiguration(false))
                {
                    ShowSettings();
                }

                Application.Run(window);
            }
            finally
            {
                notifyIcon.Visible = false;
                notifyIcon.Dispose();
                window.Dispose();
                iconResources.Dispose();
            }
        }

        private void ConfigureTrayMenu()
        {
            var menu = new ContextMenuStrip();

            var nextItem = menu.Items.Add("Switch To Next Output");
            nextItem.Click += delegate { SwitchToNextDevice(); };

            var settingsItem = menu.Items.Add("Settings");
            settingsItem.Click += delegate { ShowSettings(); };

            menu.Items.Add("-");

            var exitItem = menu.Items.Add("Exit");
            exitItem.Click += delegate { window.Close(); };

            notifyIcon.ContextMenuStrip = menu;
            notifyIcon.DoubleClick += delegate { ShowSettings(); };
        }

        private bool ApplyConfiguration(bool showSuccessBalloon)
        {
            try
            {
                StartupManager.Apply(config);
                configuredDevices = ResolveConfiguredDevices(config);

                if (configuredDevices.Count < 2)
                {
                    notifyIcon.Text = "Audio Switcher - configuration needed";
                    ShowConfigurationWarning("Please select at least two output devices in Settings.\n\nAvailable devices:\n" + BuildAvailableDeviceList());
                    return false;
                }

                currentDeviceIndex = FindInitialDeviceIndex(configuredDevices);

                if (!RegisterHotkey())
                {
                    notifyIcon.Text = "Audio Switcher - hotkey conflict";
                    ShowConfigurationWarning("Hotkey registration failed. Pick another hotkey in Settings.");
                    return false;
                }

                notifyIcon.Text = "Audio Switcher";

                if (showSuccessBalloon)
                {
                    notifyIcon.ShowBalloonTip(1200, "Audio Switcher", "Settings saved. Hotkey is now active.", ToolTipIcon.Info);
                }

                return true;
            }
            catch
            {
                notifyIcon.Text = "Audio Switcher - configuration error";
                ShowConfigurationWarning("Configuration could not be applied.");
                return false;
            }
        }

        private bool RegisterHotkey()
        {
            window.UnregisterAll();

            if (config.Hotkey == null || string.IsNullOrWhiteSpace(config.Hotkey.Key))
            {
                return false;
            }

            try
            {
                var modifiers = ParseModifiers(config.Hotkey.Modifiers);
                var key = (Keys)Enum.Parse(typeof(Keys), config.Hotkey.Key, true);
                window.Register(1, modifiers, key);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ShowSettings()
        {
            var devices = AudioDeviceManager.GetRenderDevices();

            using (var form = new SettingsForm(config, devices))
            {
                if (form.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                config = form.BuildConfig();
                SaveConfig(configPath, config);
                ApplyConfiguration(true);
            }
        }

        private void SwitchToNextDevice()
        {
            if (configuredDevices == null || configuredDevices.Count == 0)
            {
                ShowConfigurationWarning("No switchable devices are configured. Open Settings from the tray icon.");
                return;
            }

            currentDeviceIndex = (currentDeviceIndex + 1) % configuredDevices.Count;
            var device = configuredDevices[currentDeviceIndex];
            AudioDeviceManager.SetDefaultRenderDevice(device.Id);
            ShowSwitchNotification(device.Name);
        }

        private void ShowSwitchNotification(string deviceName)
        {
            var notifications = config.Notifications ?? NotificationConfig.CreateDefault();

            if (notifications.Tray)
            {
                notifyIcon.ShowBalloonTip(1200, "Audio Output Switched", deviceName, ToolTipIcon.Info);
            }

            if (notifications.Overlay)
            {
                OverlayForm.ShowMessage(deviceName, notifications.OverlayDurationMs);
            }
        }

        private static int FindInitialDeviceIndex(List<AudioDeviceInfo> devices)
        {
            for (var i = 0; i < devices.Count; i++)
            {
                if (devices[i].IsDefault)
                {
                    return i;
                }
            }

            return 0;
        }

        private static AppConfig LoadOrCreateConfig(string path)
        {
            if (!File.Exists(path))
            {
                var defaultConfig = AppConfig.CreateDefault();
                SaveConfig(path, defaultConfig);
                return defaultConfig;
            }

            var json = File.ReadAllText(path, Encoding.UTF8);
            var serializer = new JavaScriptSerializer();
            return AppConfig.Normalize(serializer.Deserialize<AppConfig>(json));
        }

        private static void SaveConfig(string path, AppConfig appConfig)
        {
            var serializer = new JavaScriptSerializer();
            var json = serializer.Serialize(AppConfig.Normalize(appConfig));
            File.WriteAllText(path, json, new UTF8Encoding(true));
        }

        private static List<AudioDeviceInfo> ResolveConfiguredDevices(AppConfig loadedConfig)
        {
            var allDevices = AudioDeviceManager.GetRenderDevices();
            var result = new List<AudioDeviceInfo>();

            if (loadedConfig.Devices == null)
            {
                return result;
            }

            foreach (var entry in loadedConfig.Devices)
            {
                if (entry == null)
                {
                    continue;
                }

                AudioDeviceInfo matchedDevice = null;

                if (!string.IsNullOrWhiteSpace(entry.DeviceId))
                {
                    foreach (var device in allDevices)
                    {
                        if (string.Equals(device.Id, entry.DeviceId, StringComparison.OrdinalIgnoreCase))
                        {
                            matchedDevice = device;
                            break;
                        }
                    }
                }

                if (matchedDevice == null && !string.IsNullOrWhiteSpace(entry.Match))
                {
                    foreach (var device in allDevices)
                    {
                        if (device.Name.IndexOf(entry.Match, StringComparison.CurrentCultureIgnoreCase) >= 0)
                        {
                            matchedDevice = device;
                            break;
                        }
                    }
                }

                if (matchedDevice != null && !ContainsDevice(result, matchedDevice.Id))
                {
                    result.Add(matchedDevice);
                }
            }

            return result;
        }

        private static bool ContainsDevice(List<AudioDeviceInfo> devices, string id)
        {
            foreach (var device in devices)
            {
                if (string.Equals(device.Id, id, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static HotkeyModifiers ParseModifiers(IEnumerable<string> modifiers)
        {
            HotkeyModifiers result = HotkeyModifiers.None;

            if (modifiers == null)
            {
                return result;
            }

            foreach (var modifier in modifiers)
            {
                result |= (HotkeyModifiers)Enum.Parse(typeof(HotkeyModifiers), modifier, true);
            }

            return result;
        }

        private string BuildAvailableDeviceList()
        {
            var builder = new StringBuilder();
            var devices = AudioDeviceManager.GetRenderDevices();

            foreach (var device in devices)
            {
                builder.Append(device.IsDefault ? "* " : "  ");
                builder.AppendLine(device.Name);
            }

            return builder.ToString();
        }

        private static void ShowConfigurationWarning(string message)
        {
            MessageBox.Show(message, "Audio Switcher", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    internal sealed class SettingsForm : Form
    {
        private readonly CheckedListBox deviceList;
        private readonly CheckBox controlModifier;
        private readonly CheckBox shiftModifier;
        private readonly CheckBox altModifier;
        private readonly CheckBox winModifier;
        private readonly ComboBox keyComboBox;
        private readonly CheckBox runAtStartup;
        private readonly CheckBox trayNotification;
        private readonly CheckBox overlayNotification;
        private readonly NumericUpDown overlayDuration;
        private readonly List<AudioDeviceInfo> devices;
        private readonly AppConfig config;

        public SettingsForm(AppConfig currentConfig, List<AudioDeviceInfo> availableDevices)
        {
            config = AppConfig.Normalize(currentConfig);
            devices = availableDevices;

            Text = "Audio Switcher Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 520);

            var deviceLabel = new Label { Left = 16, Top = 16, Width = 470, Text = "Select the output devices to cycle through:" };
            Controls.Add(deviceLabel);

            deviceList = new CheckedListBox { Left = 16, Top = 44, Width = 486, Height = 180, CheckOnClick = true };
            Controls.Add(deviceList);

            var hotkeyLabel = new Label { Left = 16, Top = 238, Width = 120, Text = "Hotkey:" };
            Controls.Add(hotkeyLabel);

            controlModifier = new CheckBox { Left = 16, Top = 266, Width = 80, Text = "Ctrl" };
            shiftModifier = new CheckBox { Left = 96, Top = 266, Width = 80, Text = "Shift" };
            altModifier = new CheckBox { Left = 176, Top = 266, Width = 80, Text = "Alt" };
            winModifier = new CheckBox { Left = 256, Top = 266, Width = 80, Text = "Win" };
            Controls.Add(controlModifier);
            Controls.Add(shiftModifier);
            Controls.Add(altModifier);
            Controls.Add(winModifier);

            keyComboBox = new ComboBox { Left = 352, Top = 263, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            Controls.Add(keyComboBox);

            var startupLabel = new Label { Left = 16, Top = 316, Width = 160, Text = "Startup:" };
            Controls.Add(startupLabel);

            runAtStartup = new CheckBox { Left = 16, Top = 344, Width = 240, Text = "Launch automatically at sign-in" };
            Controls.Add(runAtStartup);

            var notificationLabel = new Label { Left = 16, Top = 382, Width = 160, Text = "Notifications:" };
            Controls.Add(notificationLabel);

            trayNotification = new CheckBox { Left = 16, Top = 410, Width = 220, Text = "Tray balloon notification" };
            overlayNotification = new CheckBox { Left = 16, Top = 438, Width = 240, Text = "On-screen overlay notification" };
            Controls.Add(trayNotification);
            Controls.Add(overlayNotification);

            var durationLabel = new Label { Left = 16, Top = 472, Width = 180, Text = "Overlay duration (ms):" };
            Controls.Add(durationLabel);

            overlayDuration = new NumericUpDown { Left = 200, Top = 468, Width = 100, Minimum = 500, Maximum = 5000, Increment = 100 };
            Controls.Add(overlayDuration);

            var saveButton = new Button { Left = 336, Top = 460, Width = 80, Text = "Save" };
            saveButton.Click += SaveButtonOnClick;
            Controls.Add(saveButton);

            var cancelButton = new Button { Left = 422, Top = 460, Width = 80, Text = "Cancel" };
            cancelButton.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(cancelButton);

            AcceptButton = saveButton;
            CancelButton = cancelButton;

            PopulateKeyChoices();
            PopulateDeviceChoices();
            ApplyCurrentConfigToControls();
        }

        public AppConfig BuildConfig()
        {
            var nextConfig = AppConfig.Normalize(config);
            nextConfig.Hotkey = new HotkeyConfig { Modifiers = BuildModifierList(), Key = (string)keyComboBox.SelectedItem };
            nextConfig.Devices = new List<DeviceMatchConfig>();
            nextConfig.Startup = new StartupConfig { RunAtLogin = runAtStartup.Checked };

            for (var index = 0; index < deviceList.Items.Count; index++)
            {
                if (deviceList.GetItemChecked(index))
                {
                    var device = devices[index];
                    nextConfig.Devices.Add(new DeviceMatchConfig { DeviceId = device.Id, Match = device.Name, DisplayName = device.Name });
                }
            }

            nextConfig.Notifications = new NotificationConfig
            {
                Tray = trayNotification.Checked,
                Overlay = overlayNotification.Checked,
                OverlayDurationMs = (int)overlayDuration.Value
            };

            return nextConfig;
        }

        private void PopulateKeyChoices()
        {
            for (var i = 1; i <= 12; i++)
            {
                keyComboBox.Items.Add("F" + i);
            }

            for (var i = 'A'; i <= 'Z'; i++)
            {
                keyComboBox.Items.Add(((char)i).ToString());
            }

            for (var i = 0; i <= 9; i++)
            {
                keyComboBox.Items.Add("D" + i);
            }
        }

        private void PopulateDeviceChoices()
        {
            foreach (var device in devices)
            {
                deviceList.Items.Add(device.Name, false);
            }
        }

        private void ApplyCurrentConfigToControls()
        {
            if (config.Hotkey != null && config.Hotkey.Modifiers != null)
            {
                foreach (var modifier in config.Hotkey.Modifiers)
                {
                    if (string.Equals(modifier, "Control", StringComparison.OrdinalIgnoreCase)) controlModifier.Checked = true;
                    if (string.Equals(modifier, "Shift", StringComparison.OrdinalIgnoreCase)) shiftModifier.Checked = true;
                    if (string.Equals(modifier, "Alt", StringComparison.OrdinalIgnoreCase)) altModifier.Checked = true;
                    if (string.Equals(modifier, "Win", StringComparison.OrdinalIgnoreCase)) winModifier.Checked = true;
                }
            }

            if (config.Hotkey != null && !string.IsNullOrWhiteSpace(config.Hotkey.Key))
            {
                keyComboBox.SelectedItem = config.Hotkey.Key;
            }

            if (keyComboBox.SelectedIndex < 0 && keyComboBox.Items.Count > 0)
            {
                keyComboBox.SelectedIndex = 0;
            }

            for (var i = 0; i < devices.Count; i++)
            {
                if (IsConfiguredDevice(devices[i]))
                {
                    deviceList.SetItemChecked(i, true);
                }
            }

            var startup = config.Startup ?? StartupConfig.CreateDefault();
            runAtStartup.Checked = startup.RunAtLogin;

            var notifications = config.Notifications ?? NotificationConfig.CreateDefault();
            trayNotification.Checked = notifications.Tray;
            overlayNotification.Checked = notifications.Overlay;
            overlayDuration.Value = Math.Max(overlayDuration.Minimum, Math.Min(overlayDuration.Maximum, notifications.OverlayDurationMs));
        }

        private bool IsConfiguredDevice(AudioDeviceInfo device)
        {
            if (config.Devices == null)
            {
                return false;
            }

            foreach (var configuredDevice in config.Devices)
            {
                if (configuredDevice == null) continue;

                if (!string.IsNullOrWhiteSpace(configuredDevice.DeviceId) &&
                    string.Equals(configuredDevice.DeviceId, device.Id, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(configuredDevice.Match) &&
                    device.Name.IndexOf(configuredDevice.Match, StringComparison.CurrentCultureIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private string[] BuildModifierList()
        {
            var result = new List<string>();
            if (controlModifier.Checked) result.Add("Control");
            if (shiftModifier.Checked) result.Add("Shift");
            if (altModifier.Checked) result.Add("Alt");
            if (winModifier.Checked) result.Add("Win");
            return result.ToArray();
        }

        private void SaveButtonOnClick(object sender, EventArgs e)
        {
            if (deviceList.CheckedItems.Count < 2)
            {
                MessageBox.Show("Please select at least two output devices.", "Audio Switcher", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!controlModifier.Checked && !shiftModifier.Checked && !altModifier.Checked && !winModifier.Checked)
            {
                MessageBox.Show("Please choose at least one hotkey modifier.", "Audio Switcher", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (keyComboBox.SelectedItem == null)
            {
                MessageBox.Show("Please choose a hotkey key.", "Audio Switcher", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }

    internal sealed class AppIconResources : IDisposable
    {
        private readonly Bitmap bitmap;
        private readonly IntPtr iconHandle;

        public Icon Icon { get; private set; }

        private AppIconResources(Bitmap bitmap, IntPtr iconHandle, Icon icon)
        {
            this.bitmap = bitmap;
            this.iconHandle = iconHandle;
            Icon = icon;
        }

        public static AppIconResources Create(string baseDirectory)
        {
            var iconPath = Path.Combine(baseDirectory, "assets", "app-icon.ico");
            if (File.Exists(iconPath))
            {
                return new AppIconResources(null, IntPtr.Zero, new Icon(iconPath));
            }

            var bitmap = new Bitmap(16, 16);
            using (var graphics = Graphics.FromImage(bitmap))
            using (var brush = new SolidBrush(Color.FromArgb(0, 120, 215)))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillEllipse(brush, 0, 0, 15, 15);
            }

            var iconHandle = bitmap.GetHicon();
            var icon = Icon.FromHandle(iconHandle);
            return new AppIconResources(bitmap, iconHandle, icon);
        }

        public void Dispose()
        {
            if (Icon != null)
            {
                Icon.Dispose();
                Icon = null;
            }

            if (iconHandle != IntPtr.Zero)
            {
                DestroyIcon(iconHandle);
            }

            if (bitmap != null)
            {
                bitmap.Dispose();
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr handle);
    }

    internal sealed class OverlayForm : Form
    {
        private readonly Timer timer;

        private OverlayForm(string message, int durationMs)
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            ShowInTaskbar = false;
            BackColor = Color.FromArgb(30, 30, 30);
            Opacity = 0.9;
            Width = 420;
            Height = 86;

            var area = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(area.Right - Width - 24, area.Bottom - Height - 24);

            var label = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.White,
                Text = message
            };

            Controls.Add(label);

            timer = new Timer { Interval = durationMs };
            timer.Tick += delegate
            {
                timer.Stop();
                Close();
            };
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            timer.Start();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && timer != null)
            {
                timer.Dispose();
            }

            base.Dispose(disposing);
        }

        public static void ShowMessage(string message, int durationMs)
        {
            var form = new OverlayForm(message, durationMs);
            form.Show();
        }
    }

    [Flags]
    internal enum HotkeyModifiers
    {
        None = 0x0000,
        Alt = 0x0001,
        Control = 0x0002,
        Shift = 0x0004,
        Win = 0x0008
    }

    internal sealed class HotkeyWindow : Form
    {
        private const int WmHotkey = 0x0312;
        private readonly List<int> registeredIds = new List<int>();

        public event EventHandler HotkeyPressed;

        public HotkeyWindow()
        {
            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            WindowState = FormWindowState.Minimized;
            Opacity = 0;
            Width = 0;
            Height = 0;
        }

        public void Register(int id, HotkeyModifiers modifiers, Keys key)
        {
            if (!RegisterHotKey(Handle, id, (uint)modifiers, (uint)key))
            {
                throw new InvalidOperationException("RegisterHotKey failed.");
            }

            registeredIds.Add(id);
        }

        public void UnregisterAll()
        {
            foreach (var id in registeredIds)
            {
                UnregisterHotKey(Handle, id);
            }

            registeredIds.Clear();
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(false);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmHotkey && HotkeyPressed != null)
            {
                HotkeyPressed(this, EventArgs.Empty);
            }

            base.WndProc(ref m);
        }

        protected override void Dispose(bool disposing)
        {
            UnregisterAll();
            base.Dispose(disposing);
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }

    internal sealed class AppConfig
    {
        public HotkeyConfig Hotkey { get; set; }
        public List<DeviceMatchConfig> Devices { get; set; }
        public StartupConfig Startup { get; set; }
        public NotificationConfig Notifications { get; set; }

        public static AppConfig CreateDefault()
        {
            return new AppConfig
            {
                Hotkey = new HotkeyConfig
                {
                    Modifiers = new[] { "Control", "Shift" },
                    Key = "F11"
                },
                Devices = new List<DeviceMatchConfig>(),
                Startup = StartupConfig.CreateDefault(),
                Notifications = NotificationConfig.CreateDefault()
            };
        }

        public static AppConfig Normalize(AppConfig value)
        {
            var result = value ?? CreateDefault();

            if (result.Hotkey == null)
            {
                result.Hotkey = CreateDefault().Hotkey;
            }

            if (result.Devices == null)
            {
                result.Devices = new List<DeviceMatchConfig>();
            }

            if (result.Startup == null)
            {
                result.Startup = StartupConfig.CreateDefault();
            }

            if (result.Notifications == null)
            {
                result.Notifications = NotificationConfig.CreateDefault();
            }

            return result;
        }
    }

    internal sealed class HotkeyConfig
    {
        public string[] Modifiers { get; set; }
        public string Key { get; set; }
    }

    internal sealed class DeviceMatchConfig
    {
        public string DeviceId { get; set; }
        public string Match { get; set; }
        public string DisplayName { get; set; }
    }

    internal sealed class StartupConfig
    {
        public bool RunAtLogin { get; set; }

        public static StartupConfig CreateDefault()
        {
            return new StartupConfig { RunAtLogin = false };
        }
    }

    internal sealed class NotificationConfig
    {
        public bool Tray { get; set; }
        public bool Overlay { get; set; }
        public int OverlayDurationMs { get; set; }

        public static NotificationConfig CreateDefault()
        {
            return new NotificationConfig
            {
                Tray = true,
                Overlay = true,
                OverlayDurationMs = 1500
            };
        }
    }

    internal static class StartupManager
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "AudioSwitcher";

        public static void Apply(AppConfig config)
        {
            var startup = config.Startup ?? StartupConfig.CreateDefault();
            using (var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, true))
            {
                if (key == null)
                {
                    return;
                }

                if (startup.RunAtLogin)
                {
                    var exePath = Application.ExecutablePath;
                    key.SetValue(AppName, "\"" + exePath + "\"", RegistryValueKind.String);
                }
                else
                {
                    key.DeleteValue(AppName, false);
                }
            }
        }
    }

    internal sealed class AudioDeviceInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public bool IsDefault { get; set; }
    }

    internal enum EDataFlow
    {
        ERender,
        ECapture,
        EAll,
        EDataFlowEnumCount
    }

    internal enum ERole
    {
        EConsole,
        EMultimedia,
        ECommunications,
        ERoleEnumCount
    }

    [Flags]
    internal enum DeviceState : uint
    {
        Active = 0x1,
        Disabled = 0x2,
        NotPresent = 0x4,
        Unplugged = 0x8,
        All = 0xF
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PropertyKey
    {
        public Guid FormatId;
        public int PropertyId;

        public PropertyKey(Guid formatId, int propertyId)
        {
            FormatId = formatId;
            PropertyId = propertyId;
        }
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct PropVariant
    {
        [FieldOffset(0)]
        public ushort ValueType;

        [FieldOffset(8)]
        public IntPtr PointerValue;

        public string GetValue()
        {
            if (ValueType == 31 && PointerValue != IntPtr.Zero)
            {
                return Marshal.PtrToStringUni(PointerValue);
            }

            return null;
        }
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumeratorComObject
    {
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(EDataFlow dataFlow, DeviceState stateMask, out IMMDeviceCollection devices);
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice endpoint);
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string id, out IMMDevice device);
        int RegisterEndpointNotificationCallback(IntPtr client);
        int UnregisterEndpointNotificationCallback(IntPtr client);
    }

    [ComImport]
    [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceCollection
    {
        int GetCount(out uint count);
        int Item(uint index, out IMMDevice device);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        int Activate(ref Guid iid, uint clsCtx, IntPtr activationParams, out IntPtr interfacePointer);
        int OpenPropertyStore(uint stgmAccess, out IPropertyStore properties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetState(out DeviceState state);
    }

    [ComImport]
    [Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyStore
    {
        int GetCount(out uint propertyCount);
        int GetAt(uint propertyIndex, out PropertyKey key);
        int GetValue(ref PropertyKey key, out PropVariant value);
        int SetValue(ref PropertyKey key, ref PropVariant value);
        int Commit();
    }

    [ComImport]
    [Guid("F8679F50-850A-41CF-9C72-430F290290C8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPolicyConfig
    {
        int GetMixFormat();
        int GetDeviceFormat();
        int ResetDeviceFormat();
        int SetDeviceFormat();
        int GetProcessingPeriod();
        int SetProcessingPeriod();
        int GetShareMode();
        int SetShareMode();
        int GetPropertyValue();
        int SetPropertyValue();
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string deviceId, ERole role);
        int SetEndpointVisibility();
    }

    [ComImport]
    [Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
    internal class PolicyConfigClient
    {
    }

    internal static class AudioDeviceManager
    {
        private static readonly PropertyKey FriendlyNameKey =
            new PropertyKey(new Guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), 14);

        public static List<AudioDeviceInfo> GetRenderDevices()
        {
            var devices = new List<AudioDeviceInfo>();
            var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();

            IMMDeviceCollection collection;
            Marshal.ThrowExceptionForHR(enumerator.EnumAudioEndpoints(EDataFlow.ERender, DeviceState.Active, out collection));

            IMMDevice defaultDevice;
            Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(EDataFlow.ERender, ERole.EMultimedia, out defaultDevice));

            string defaultId;
            Marshal.ThrowExceptionForHR(defaultDevice.GetId(out defaultId));

            uint count;
            Marshal.ThrowExceptionForHR(collection.GetCount(out count));

            for (uint index = 0; index < count; index++)
            {
                IMMDevice device;
                Marshal.ThrowExceptionForHR(collection.Item(index, out device));

                string id;
                Marshal.ThrowExceptionForHR(device.GetId(out id));

                IPropertyStore propertyStore;
                Marshal.ThrowExceptionForHR(device.OpenPropertyStore(0, out propertyStore));

                PropVariant value;
                var key = FriendlyNameKey;
                Marshal.ThrowExceptionForHR(propertyStore.GetValue(ref key, out value));

                devices.Add(new AudioDeviceInfo
                {
                    Id = id,
                    Name = value.GetValue() ?? id,
                    IsDefault = string.Equals(id, defaultId, StringComparison.OrdinalIgnoreCase)
                });
            }

            return devices;
        }

        public static void SetDefaultRenderDevice(string deviceId)
        {
            var policyConfig = (IPolicyConfig)new PolicyConfigClient();
            Marshal.ThrowExceptionForHR(policyConfig.SetDefaultEndpoint(deviceId, ERole.EConsole));
            Marshal.ThrowExceptionForHR(policyConfig.SetDefaultEndpoint(deviceId, ERole.EMultimedia));
            Marshal.ThrowExceptionForHR(policyConfig.SetDefaultEndpoint(deviceId, ERole.ECommunications));
        }
    }
}
