# Audio Switcher

Prompt: 一个简单的Windows应用，支持通过自定义热键来切换系统默认的音频输出选项，可自行配置在哪些输出选项中切换，切换时在状态栏或者屏幕上有提示。输出选择和热键设置都要做成gui界面，在状态栏有图标显示。

A lightweight Windows tray utility for switching between selected audio output devices with a custom global hotkey.

![Audio Switcher Logo](assets/app-logo.png)

## Overview

Audio Switcher lives in the Windows system tray and lets you:

- choose which output devices are part of the switch cycle
- configure a global hotkey from a GUI settings window
- switch the default audio output instantly
- see tray and on-screen notifications when the output changes

This project is built as a native Windows `.exe` using C# WinForms and the Windows Core Audio APIs.

## Features

- Tray icon with context menu
- GUI settings window
- Custom global hotkey
- Selectable output device list
- Tray balloon notifications
- On-screen overlay notifications
- Configuration saved to `config.json`
- Standalone Windows executable

## Demo Flow

1. Launch `AudioSwitcher.exe`
2. Open `Settings` from the tray icon
3. Select at least two output devices
4. Choose a hotkey combination
5. Save
6. Press the hotkey to cycle to the next configured output

## Project Files

- `AudioSwitcher.cs` - main WinForms application source
- `AudioSwitcher.exe` - compiled app
- `config.json` - user configuration
- `build-exe.bat` - local build script
- `audio-switcher.ps1` - earlier PowerShell prototype
- `start-audio-switcher.bat` - launcher for the PowerShell prototype

## Build

This project uses the classic .NET Framework compiler available on Windows.

```bat
build-exe.bat
```

That generates:

```text
AudioSwitcher.exe
```

## Run

```bat
AudioSwitcher.exe
```

After launch:

- the app appears in the Windows notification area
- double-click the tray icon to open settings
- right-click the tray icon for the full menu
- use your configured hotkey to switch outputs

## Configuration

The application stores settings in `config.json` next to the executable.

The GUI can configure:

- selected audio output devices
- hotkey modifiers
- hotkey key
- tray notifications
- overlay notifications
- overlay duration


## Tech Notes

- Uses Windows Core Audio device enumeration
- Sets the default render endpoint for console, multimedia, and communications roles
- Stores selected devices by device ID for more reliable matching

## Publishing To GitHub

Basic local publish flow:

```bat
git add .
git commit -m "Initial release"
git remote add origin https://github.com/YOUR_NAME/audio-switcher.git
git push -u origin master
```

## License

MIT
