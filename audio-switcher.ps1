[CmdletBinding()]
param(
    [string]$ConfigPath = "config.json",
    [switch]$ListDevices
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not [System.IO.Path]::IsPathRooted($ConfigPath)) {
    $ConfigPath = Join-Path $scriptRoot $ConfigPath
}

$audioInteropSource = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AudioSwitcher.Interop
{
    public enum EDataFlow
    {
        eRender,
        eCapture,
        eAll,
        EDataFlow_enum_count
    }

    public enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
        ERole_enum_count
    }

    [Flags]
    public enum DeviceState : uint
    {
        Active = 0x1,
        Disabled = 0x2,
        NotPresent = 0x4,
        Unplugged = 0x8,
        All = 0xF
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public int pid;

        public PROPERTYKEY(Guid format, int propertyId)
        {
            fmtid = format;
            pid = propertyId;
        }
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct PROPVARIANT
    {
        [FieldOffset(0)]
        public ushort vt;

        [FieldOffset(8)]
        public IntPtr pointerValue;

        public string GetValue()
        {
            if (vt == 31 && pointerValue != IntPtr.Zero)
            {
                return Marshal.PtrToStringUni(pointerValue);
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
        int GetDevice(string id, out IMMDevice device);
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
        int GetAt(uint propertyIndex, out PROPERTYKEY key);
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT value);
        int SetValue(ref PROPERTYKEY key, ref PROPVARIANT value);
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
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string wszDeviceId, ERole role);
        int SetEndpointVisibility();
    }

    [ComImport]
    [Guid("870af99c-171d-4f9e-af0d-e63df40c2bc9")]
    internal class PolicyConfigClient
    {
    }

    public sealed class AudioDeviceInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public bool IsDefault { get; set; }
    }

    public static class AudioDeviceManager
    {
        private static readonly PROPERTYKEY FriendlyNameKey =
            new PROPERTYKEY(new Guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), 14);

        public static List<AudioDeviceInfo> GetRenderDevices()
        {
            var devices = new List<AudioDeviceInfo>();
            IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
            IMMDeviceCollection collection;
            Marshal.ThrowExceptionForHR(enumerator.EnumAudioEndpoints(EDataFlow.eRender, DeviceState.Active, out collection));

            IMMDevice defaultDevice;
            Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out defaultDevice));
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

                IPropertyStore store;
                Marshal.ThrowExceptionForHR(device.OpenPropertyStore(0, out store));

                PROPVARIANT value;
                var key = FriendlyNameKey;
                Marshal.ThrowExceptionForHR(store.GetValue(ref key, out value));

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
            Marshal.ThrowExceptionForHR(policyConfig.SetDefaultEndpoint(deviceId, ERole.eConsole));
            Marshal.ThrowExceptionForHR(policyConfig.SetDefaultEndpoint(deviceId, ERole.eMultimedia));
            Marshal.ThrowExceptionForHR(policyConfig.SetDefaultEndpoint(deviceId, ERole.eCommunications));
        }
    }
}
"@

$windowInteropSource = @"
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace AudioSwitcher.Interop
{
    public delegate void HotkeyPressedHandler(object sender, int hotkeyId);

    [Flags]
    public enum HotkeyModifiers
    {
        None = 0x0000,
        Alt = 0x0001,
        Control = 0x0002,
        Shift = 0x0004,
        Win = 0x0008
    }

    public sealed class HotkeyWindow : Form
    {
        private const int WM_HOTKEY = 0x0312;
        private readonly List<int> registeredIds = new List<int>();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public event HotkeyPressedHandler HotkeyPressed;

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

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(false);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                if (HotkeyPressed != null)
                {
                    HotkeyPressed(this, m.WParam.ToInt32());
                }
            }

            base.WndProc(ref m);
        }

        protected override void Dispose(bool disposing)
        {
            foreach (var id in registeredIds)
            {
                UnregisterHotKey(Handle, id);
            }

            registeredIds.Clear();
            base.Dispose(disposing);
        }
    }
}
"@

Add-Type -TypeDefinition $audioInteropSource
Add-Type -TypeDefinition $windowInteropSource -ReferencedAssemblies @(
    "System.dll",
    "System.Windows.Forms.dll"
)

function Get-DefaultConfig {
    @{
        hotkey = @{
            modifiers = @("Control", "Alt")
            key = "F10"
        }
        devices = @(
            @{ match = "Headphones" },
            @{ match = "Speakers" }
        )
        notifications = @{
            tray = $true
            overlay = $true
            overlayDurationMs = 1500
        }
    }
}

function Save-JsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        $Value
    )

    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }

    $json = $Value | ConvertTo-Json -Depth 8
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $json, $utf8Bom)
}

function Ensure-Config {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Save-JsonFile -Path $Path -Value (Get-DefaultConfig)
        Write-Host "Config file created: $Path"
        Write-Host "Update devices / hotkey first, then run again."
        exit 0
    }

    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
}

function Get-HotkeyModifiersValue {
    param([object[]]$Modifiers)

    $result = [AudioSwitcher.Interop.HotkeyModifiers]::None
    foreach ($item in $Modifiers) {
        $name = [string]$item
        $result = $result -bor ([AudioSwitcher.Interop.HotkeyModifiers]::$name)
    }

    return $result
}

function Get-ConfiguredDeviceList {
    param(
        [object[]]$ConfiguredDevices,
        [System.Collections.Generic.List[AudioSwitcher.Interop.AudioDeviceInfo]]$AvailableDevices
    )

    $selected = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $ConfiguredDevices) {
        if ($null -eq $entry.match -or [string]::IsNullOrWhiteSpace([string]$entry.match)) {
            continue
        }

        $matched = $AvailableDevices | Where-Object { $_.Name -like "*$($entry.match)*" }
        foreach ($device in $matched) {
            if (-not ($selected | Where-Object { $_.Id -eq $device.Id })) {
                $selected.Add($device)
            }
        }
    }

    return $selected.ToArray()
}

function Show-AvailableDevices {
    param([System.Collections.Generic.List[AudioSwitcher.Interop.AudioDeviceInfo]]$Devices)

    Write-Host "Available output devices:"
    foreach ($device in $Devices) {
        $marker = if ($device.IsDefault) { "*" } else { " " }
        Write-Host ("{0} {1}" -f $marker, $device.Name)
    }
}

function Show-Overlay {
    param(
        [string]$Message,
        [int]$DurationMs = 1500
    )

    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $form.Opacity = 0.9
    $form.Width = 420
    $form.Height = 86

    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $form.Location = New-Object System.Drawing.Point(
        ($screen.Right - $form.Width - 24),
        ($screen.Bottom - $form.Height - 24)
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Dock = [System.Windows.Forms.DockStyle]::Fill
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::White
    $label.Text = $Message
    $form.Controls.Add($label)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $DurationMs
    $timer.Add_Tick({
        $timer.Stop()
        $form.Close()
        $form.Dispose()
        $timer.Dispose()
    })

    $timer.Start()
    $form.Show()
}

function New-StatusIcon {
    $bitmap = New-Object System.Drawing.Bitmap 16, 16
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 120, 215))
    $graphics.FillEllipse($brush, 0, 0, 15, 15)
    $graphics.Dispose()
    $brush.Dispose()

    $icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
    return @{
        Bitmap = $bitmap
        Icon = $icon
    }
}

$config = Ensure-Config -Path $ConfigPath
$allDevices = [AudioSwitcher.Interop.AudioDeviceManager]::GetRenderDevices()

if ($ListDevices) {
    Show-AvailableDevices -Devices $allDevices
    exit 0
}

$configuredDevices = @(Get-ConfiguredDeviceList -ConfiguredDevices $config.devices -AvailableDevices $allDevices)

if ($configuredDevices.Count -lt 2) {
    Show-AvailableDevices -Devices $allDevices
    Write-Error "Fewer than 2 output devices matched config.json devices.match."
}

$state = [ordered]@{
    DeviceIndex = 0
    Devices = $configuredDevices
    Notifications = $config.notifications
}

for ($index = 0; $index -lt $state.Devices.Count; $index++) {
    if ($state.Devices[$index].IsDefault) {
        $state.DeviceIndex = $index
        break
    }
}

$notifyResources = New-StatusIcon
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = $notifyResources.Icon
$notifyIcon.Text = "Audio Switcher"
$notifyIcon.Visible = $true

function Show-Message {
    param([string]$Text)

    if ($state.Notifications.tray) {
        $notifyIcon.BalloonTipTitle = "Audio Output Switched"
        $notifyIcon.BalloonTipText = $Text
        $notifyIcon.ShowBalloonTip(1200)
    }

    if ($state.Notifications.overlay) {
        Show-Overlay -Message $Text -DurationMs ([int]$state.Notifications.overlayDurationMs)
    }
}

function Switch-ToNextDevice {
    $state.DeviceIndex = ($state.DeviceIndex + 1) % $state.Devices.Count
    $device = $state.Devices[$state.DeviceIndex]
    [AudioSwitcher.Interop.AudioDeviceManager]::SetDefaultRenderDevice($device.Id)
    Show-Message -Text $device.Name
}

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$switchItem = $contextMenu.Items.Add("Switch To Next Output")
$switchItem.Add_Click({ Switch-ToNextDevice })
$contextMenu.Items.Add("-") | Out-Null
$exitItem = $contextMenu.Items.Add("Exit")
$notifyIcon.ContextMenuStrip = $contextMenu

$window = New-Object AudioSwitcher.Interop.HotkeyWindow
$window.CreateControl()
$window.add_HotkeyPressed({
    param($sender, $hotkeyId)
    Switch-ToNextDevice
})

$exitItem.Add_Click({
    $window.Close()
})

$notifyIcon.Add_DoubleClick({
    Switch-ToNextDevice
})

$modifiers = Get-HotkeyModifiersValue -Modifiers $config.hotkey.modifiers
$key = [System.Windows.Forms.Keys]::$($config.hotkey.key)

try {
    $window.Register(1, $modifiers, $key)
} catch {
    $notifyIcon.Visible = $false
    throw "Global hotkey registration failed. Change config.json hotkey to avoid conflicts."
}

$notifyIcon.BalloonTipTitle = "Audio Switcher"
$notifyIcon.BalloonTipText = "Running. Hotkey: $($config.hotkey.modifiers -join '+')+$($config.hotkey.key)"
$notifyIcon.ShowBalloonTip(1500)

try {
    [System.Windows.Forms.Application]::Run($window)
} finally {
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()
    $notifyResources.Icon.Dispose()
    $notifyResources.Bitmap.Dispose()
    $window.Dispose()
}
