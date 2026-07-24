# 🦎 Chameleon

**Your Wi-Fi, adapted to your surroundings.**

Chameleon watches for a specific device to attach to your Windows 11 machine —
an external monitor, a USB-C dock, anything that shows up in Device Manager —
and automatically switches your Wi-Fi to the network you want for that spot. Dock
at the office and it jumps to `OfficeWiFi`; unplug and it drops back to `HomeWiFi`.

Windows has no built-in setting for this. Chameleon is a small, dependency-free
PowerShell tool that fills the gap. It comes in three flavours:

| Script | What it is | Best for |
|---|---|---|
| `Chameleon.ps1` | Background watcher, installed via Task Scheduler | Set-and-forget on your own PC |
| `Chameleon-Portable.ps1` | Same watcher, runs from a USB stick, no install | Carrying between machines |
| `Chameleon-Tray.ps1` | System-tray app with a status dot & menu | Wanting a visible on/off control |

Plus a helper, `Get-AttachedDevices.ps1`, to find the device name to key off of.

---

## How it works

`netsh wlan` can switch between Wi-Fi profiles you've **already saved**. Chameleon
watches for your target device via a WMI/CIM device event (or a light poll, in the
tray version) and calls `netsh` to connect the right profile when the device comes
or goes.

> **Prerequisite:** connect to both Wi-Fi networks manually at least once so their
> profiles are saved. Chameleon switches between known networks — it never adds new
> ones or handles passwords.

---

## 1. Find your device

Run the helper to list attached devices, grouped by class and indented:

```powershell
powershell -ExecutionPolicy Bypass -File .\Get-AttachedDevices.ps1
```

Narrow it down and show hardware IDs:

```powershell
.\Get-AttachedDevices.ps1 -Filter *monitor* -ShowId
```

Copy a device **Name** and wrap it in wildcards, e.g. `*Dell U2720Q*`. That becomes
your `$targetDeviceMatch`.

> **Tip:** A monitor plugged straight into HDMI/DisplayPort shows up as a
> `Monitor`-class device, but reconnects can be flaky if Windows treats it as the
> same known device. If you use a dock, keying off the **dock's USB device** is
> usually more reliable — it cleanly enumerates and de-enumerates on cable connect.

Confirm your Wi-Fi interface name too:

```powershell
netsh wlan show interfaces
```

---

## 2A. Base watcher — `Chameleon.ps1`

Edit the config block at the top:

```powershell
$targetDeviceMatch    = "*Your Monitor Or Dock Name*"
$wifiWhenConnected    = "OfficeWiFi"
$wifiWhenDisconnected = "HomeWiFi"
$wifiInterface        = "Wi-Fi"
```

> **Shared config:** if a `Chameleon.config.json` sits next to the script, its
> values override the inline block — so all three scripts can read one file. See
> [Portable](#2b-portable--chameleon-portableps1) for the JSON format. Without it,
> the inline block above is used.

Save it somewhere permanent (e.g. `C:\Scripts\Chameleon.ps1`), then register it to
run hidden at logon:

```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Scripts\Chameleon.ps1"'
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0
Register-ScheduledTask -TaskName "Chameleon" -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest
```

Logs go to `%LOCALAPPDATA%\Chameleon.log`. To remove it:
`Unregister-ScheduledTask -TaskName "Chameleon" -Confirm:$false`.

---

## 2B. Portable — `Chameleon-Portable.ps1`

Made to live on a USB stick. Nothing is written to the host machine — config and
log stay in the script's own folder.

1. Copy `Chameleon-Portable.ps1`, `Chameleon-Portable.bat`, and
   `Chameleon.config.example.json` to your USB drive.
2. Double-click `Chameleon-Portable.bat`. On first run it creates
   `Chameleon.config.json` next to the script and exits.
3. Edit `Chameleon.config.json` with your device and Wi-Fi names, then run the
   `.bat` again.

```json
{
  "deviceMatch": "*Your Monitor Or Dock Name*",
  "wifiWhenConnected": "OfficeWiFi",
  "wifiWhenDisconnected": "HomeWiFi",
  "wifiInterface": "Wi-Fi"
}
```

It runs hidden. Stop it by ending the `powershell.exe` process (Task Manager) or
logging off. Log: `Chameleon-Portable.log` beside the script.

---

## 2C. System tray — `Chameleon-Tray.ps1`

A lizard 🦎 tray icon, tinted by state:

- 🟢 **green** — target device present, on `wifiWhenConnected`
- ⚪ **grey** — device absent, on `wifiWhenDisconnected`

Right-click for the menu:

- **Auto-switch** — toggle automatic switching on/off
- **Start with Windows** — run Chameleon at logon (per-user `Run` key, no admin)
- **Force docked / undocked network** — switch either network manually
- **Settings…** — edit the device match and Wi-Fi profiles in-app (saves to
  `Chameleon.config.json`); also opens on double-click of the tray icon
- **About** — version and links
- **Open log** — view the activity log
- **Exit**

> **Start with Windows** points the logon entry at whatever you launched — the
> `.exe` if you ran the compiled build, or the `.vbs` launcher if you ran the
> script. Move the app to a permanent location *before* enabling it, since the
> saved path is absolute.

Balloon tips announce each switch. Edit the config block at the top of the script
(or drop a shared `Chameleon.config.json` beside it — it takes precedence), then
launch silently (no console window) via the VBScript shim:

```
Double-click Chameleon-Tray.vbs
```

Run it at logon: press <kbd>Win</kbd>+<kbd>R</kbd>, type `shell:startup`, and drop a
shortcut to `Chameleon-Tray.vbs` in that folder. Log: `Chameleon-Tray.log` beside
the script.

### Build a standalone .exe

Prefer a single executable over the `.ps1` + `.vbs` pair? Compile the tray app with
[ps2exe](https://github.com/MScholtes/PS2EXE):

```powershell
Install-Module ps2exe -Scope CurrentUser
powershell -ExecutionPolicy Bypass -File .\Build-Tray.ps1
```

This produces `dist\Chameleon-Tray.exe` — a no-console GUI executable with an embedded
icon. It reads `Chameleon.config.json` (or writes its log) from the folder the `.exe`
lives in, so keep the config beside it. Put a shortcut to the `.exe` in `shell:startup`
to run at logon.

> The compiled `.exe` isn't committed (SmartScreen/AV commonly flag unsigned ps2exe
> binaries). Build it yourself, or grab it from a
> [Release](https://github.com/VAROIndustries/Chameleon/releases) if one is published.

---

## Troubleshooting

- **Nothing switches** — confirm both profiles are saved (`netsh wlan show profiles`)
  and that `$wifiInterface` matches `netsh wlan show interfaces`.
- **Wrong or flaky device detection** — re-run `Get-AttachedDevices.ps1` with the
  device plugged in vs unplugged and compare; prefer the dock's USB entry over the
  monitor's EDID identity.
- **Check the log** — every switch attempt is logged (see each section above for the
  path).

## Notes & limits

- Windows 11, PowerShell 5.1+ (ships with Windows). No external modules.
- Switches between **saved** profiles only — it doesn't add networks or store
  passwords.
- `-RunLevel Highest` is used so `netsh` reliably controls the adapter.

## License

[MIT](LICENSE) © 2026 VARO Industries
