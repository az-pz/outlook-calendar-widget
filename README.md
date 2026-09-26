# Outlook Calendar Widget

Your Outlook calendar on the Windows 11 Widgets Board (<kbd>Win</kbd>+<kbd>W</kbd>).

| Small | Medium | Large |
| :---: | :---: | :---: |
| <img src="docs/images/widget-small.png" width="250" alt="Small widget showing the next meeting"> | <img src="docs/images/widget-medium.png" width="250" alt="Medium widget showing today's agenda"> | <img src="docs/images/widget-large.png" width="250" alt="Large widget showing today's agenda"> |

<sub>Screenshots use the built-in sample calendar.</sub>

## Features

- **Reads your calendar from Outlook on this PC.** It uses classic Outlook's object model, so there's no sign-in, app registration, or cloud call. Microsoft 365 via Microsoft Graph and a sample calendar are also available.
- **Three sizes:**
  - **Small** shows the next meeting with a countdown.
  - **Medium** and **Large** show an agenda: now/next highlighting, free/busy color bars, and response status.
- **Join button** for Teams and other online meetings, 15 minutes before they start.
- **Opens Outlook in one tap.** Tap an event to open it in Outlook, or tap *Open Outlook* to jump to the calendar.
- **Customize per widget:** show today, today and tomorrow, or the next 7 days, and choose whether to show declined and all-day events.
- **Companion app** (WinUI 3, Mica, light and dark themes) for choosing the calendar source, previewing the agenda, signing in to Microsoft 365, and refreshing.

## Requirements

- Windows 11 22H2 (build 22621) or later, with **Developer Mode** turned on (Settings > System > For developers).
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0).
- [Windows App Runtime 2.5](https://learn.microsoft.com/windows/apps/windows-app-sdk/downloads) for your architecture (x64 or ARM64).
- For the default calendar source: **classic Outlook for Windows** (Microsoft 365 Apps / Office), running. The new Outlook for Windows has no API that other apps can read; to use it, pick the Microsoft 365 source instead.

## Build and install

```powershell
.\scripts\deploy.ps1                          # Debug build for this PC's architecture
.\scripts\deploy.ps1 -Configuration Release   # optimized build
```

The script does the following:

1. Stops the running widget provider and the Widgets host, which lock the build output.
2. Builds the app.
3. Registers the build output as a loose-file MSIX package, so no certificate is needed.
4. Restarts the Widgets host so it picks up the change.

Then press <kbd>Win</kbd>+<kbd>W</kbd>, choose **+**, and add **Outlook Calendar**. Re-run the script after code changes; it updates the installed app in place.

Run the tests with `dotnet test --project tests\OutlookCalendarWidget.Core.Tests`.

## Calendar sources

Pick a source in the companion app (Start menu > **Outlook Calendar Widget**).

| Source | How it works | Setup |
| --- | --- | --- |
| **Outlook on this PC** (default) | Reads the default calendar of the running classic Outlook through COM. Nothing leaves the PC. The widget doesn't start Outlook in the background; if Outlook is closed, it shows an *Open Outlook* button and keeps the last agenda it read. | Keep classic Outlook running. |
| **Microsoft 365** | Calls Microsoft Graph (`/me/calendarView`, delegated `Calendars.Read`). Sign-in goes through the Windows account broker (WAM), so work, school, and personal Microsoft accounts use single sign-on. | Needs an Entra app registration (see below). |
| **Sample calendar** | A realistic demo agenda, useful for trying out the widget. | None. |

### Microsoft 365 setup

The app needs the application (client) ID of an Entra app registration. The registration must be a public client with delegated `Calendars.Read` and the WAM redirect URI. To create one with the Azure CLI:

```powershell
az login --allow-no-subscriptions
.\scripts\register-entra-app.ps1              # prints the client ID
```

Then paste the ID into the app under **Microsoft Graph connection**. Alternatively, bake it in at build time with `.\scripts\deploy.ps1 -GraphClientId <id>`.

## How it works

```
OutlookCalendarWidget.exe
 ├─ -RegisterProcessAsComServer         → COM widget provider (started by the Widgets host)
 ├─ outlook-calendar-widget:open?id=…   → opens an item in Outlook, then exits
 └─ (no arguments / other links)        → single-instance WinUI 3 companion app
```

- **Widget provider** (`src/OutlookCalendarWidget/Widgets`):
  - It's an out-of-process COM server that implements the Windows App SDK `IWidgetProvider` and `IWidgetProvider2` interfaces.
  - Cards are [Adaptive Card templates](src/OutlookCalendarWidget.Core/Templates) bound to JSON data produced by `AgendaBuilder`.
  - Refresh interval: every 2 minutes while the board is open (5 minutes when it's closed) for Outlook, and every 5 or 15 minutes for Graph.
  - Once a minute, it re-renders the card so countdowns stay current.
- **Outlook access** (`Services/OutlookDesktop.cs`):
  - Uses late-bound COM, so no Office interop assemblies are needed.
  - Calls `Items.Restrict` with recurrences expanded.
  - Reads only properties outside Outlook's object-model guard, so it never triggers the "A program is trying to access…" prompt.
  - Teams join links come from the meeting's `SkypeTeamsMeetingUrl` property.
- **Core library** (`src/OutlookCalendarWidget.Core`): pure, unit-tested logic covering agenda building, Graph and Outlook mapping, card data, and templates.
- The companion app and the widget provider share settings through the package's `LocalState` folder and signal each other through a named event.

## Troubleshooting

- **Logs:** `%LOCALAPPDATA%\Packages\OutlookCalendarWidget_tt1wz9sd5daxr\LocalState\logs\diagnostics.log` records provider starts, refresh results, and errors, but never event details.
- **The widget doesn't appear in the picker, or shows old content after a rebuild.** The Widgets host caches providers. Re-run `deploy.ps1`, or end *Windows Widgets* (`WidgetBoard.exe` and `WidgetService.exe`) in Task Manager, then press <kbd>Win</kbd>+<kbd>W</kbd> again.
- **"Couldn't connect to Outlook".** Outlook and the widget must run at the same integrity level. Don't run Outlook as administrator.
- **"Classic Outlook isn't running", but Outlook is open.** You're using the new Outlook. Switch back to classic Outlook with the toggle in its title bar, or choose the Microsoft 365 source.
- **Build error `PRI210 … 0x800704c8`.** The Widgets host has `resources.pri` open. `deploy.ps1` stops it first; if you build another way, close the Widgets board first.

## Project layout

```
src/OutlookCalendarWidget/        WinUI 3 app + widget provider (MSIX)
src/OutlookCalendarWidget.Core/   Platform-independent logic and Adaptive Card templates
tests/                            xUnit v3 tests (Microsoft Testing Platform)
scripts/                          deploy, Entra app registration, icon and picker-screenshot generators
```
