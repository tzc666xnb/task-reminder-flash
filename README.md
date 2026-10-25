# Task Reminder

**English** | [简体中文](README.zh-CN.md)

A zero-dependency, fully local task reminder for Windows that lives in the system tray. It pops up a reminder card when a task is due, keeps nagging until the task is done, rolls recurring tasks over automatically, and can start with Windows — all from a single PowerShell script. No installer, no account, no network, no telemetry.

![Main window](docs/main.png)

## Why another to-do app

- **Zero dependencies.** Windows 10/11 already ship Windows PowerShell 5.1 and WinForms. Nothing to install, nothing to update.
- **Local by design.** Every task lives in one human-readable JSON file next to the script. Copy the folder to back up; delete it to uninstall.
- **Built for "remind me at 14:30", not for project management.** Minute-precision times, time ranges, snooze, repeat rules, optional lead-time reminders.
- **Survives reboots.** A Startup-folder shortcut relaunches it silently at login.

## Screenshots

| Reminder card | Edit task |
| --- | --- |
| ![Reminder card](docs/toast.png) | ![Editor](docs/editor.png) |

## Features

| Feature | Behaviour |
| --- | --- |
| Due reminders | Polls every 10 s; a top-most card appears in the bottom-right corner, with an optional sound |
| Minute precision | Type `14:37` or spin the picker — no silent fallback to a wrong time |
| Time ranges | Optional start → end. Shows `12:30-13:30`, marks the task **in progress** until the end time, and only treats it as overdue afterwards |
| Snooze | 10 minutes from the reminder card; 10 or 30 minutes from the right-click menu |
| Lead time | Remind 5 / 10 / 15 / 30 min, 1 h, 2 h or 1 day before the due time |
| Repeats | None, daily, weekly, monthly, weekdays. Completing a recurring task schedules the next occurrence |
| Nagging | Unfinished tasks keep re-alerting every 10 / 15 / 30 / 60 minutes (default 15, can be turned off) |
| Quick add | Type a title and press `Enter` |
| Overdue visibility | Overdue rows turn red; the footer counts them |
| Autostart | A shortcut in the user's Startup folder → `wscript` → hidden PowerShell, so no console flashes |
| Single instance | A named mutex makes a second launch bring the existing window to the front |
| Crash-safe storage | Atomic replace with retries, a `.old` rollback copy, and corrupt files backed up as `tasks.json.broken-<timestamp>` |

## Requirements

- Windows 10 or 11
- Windows PowerShell 5.1 (preinstalled). No Node.js, Python, .NET SDK or administrator rights needed.

> The user interface is currently **Chinese only**. This README describes every control, and the table below maps the labels.

### UI labels

| Chinese | English |
| --- | --- |
| 任务提醒 | Task Reminder (window title) |
| 全部 / 待办 / 已完成 | All / To do / Done |
| 添加 | Add |
| 清除已完成 | Clear completed |
| 设置 | Settings |
| 不重复 / 每天 / 每周 / 每月 / 工作日 | None / Daily / Weekly / Monthly / Weekdays |
| 进行中 | In progress |
| 逾期 | Overdue |
| 稍后再提醒 | Snooze |

## Quick start

1. Download the ZIP (or `git clone`) and extract it anywhere — for example `%USERPROFILE%\Documents\TaskReminder`.
   Keep the folder together: the script locates `data\`, `app.ico` and the `.vbs` launcher relative to itself.
2. Double-click **`Start-TaskReminder.vbs`**. The main window opens and a tray icon appears.
3. *Optional* — create the Desktop and Startup shortcuts:

```powershell
powershell -ExecutionPolicy Bypass -File .\TaskReminder.ps1 -Setup
```

### Everyday use

- **Add** — type in the top box, pick date / time / repeat, press `Enter` or click 添加.
- **Complete** — tick the checkbox on the left. Clicking the row text only selects it, so you cannot complete a task by accident.
- **Edit** — double-click a row to change the time, add a note, or set a lead-time reminder.
- **Snooze / postpone** — right-click a row: snooze 10 or 30 minutes, postpone 10 minutes, postpone to tomorrow 09:00, delete.
- **Close the window** — it hides to the tray and keeps working. Quit for real from the tray icon menu.
- **Settings** — top-right 设置: sound, close-to-tray, nag interval, and start with Windows.

## Autostart

`-Setup` (or 设置 → 开机自动启动) writes
`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\任务提醒.lnk`, which runs
`wscript.exe Start-TaskReminder.vbs -StartHidden`. Using `wscript` means no console window flashes; PowerShell itself runs hidden with `-ExecutionPolicy Bypass`.

This repository ships a pre-seeded `data\settings.json` (with `"Launched": 1`), so **a fresh copy does not silently add itself to Startup**. Enable autostart explicitly in Settings, or run `-Setup`. If you delete `data\settings.json`, the app's built-in first-run behaviour enables autostart automatically.

## Files and data format

```
Task Reminder/
├─ TaskReminder.ps1          the whole app (PowerShell + WinForms, single file)
├─ Start-TaskReminder.vbs    silent launcher (no console window)
├─ app.ico                   tray / desktop icon
├─ data/
│  ├─ tasks.json             your tasks (UTF-8 JSON)
│  ├─ settings.json          preferences
│  └─ app.log                diagnostic log (safe to delete)
└─ docs/                     screenshots used by this README
```

A task looks like this:

```json
{
  "id": "7c44fa575c284f9bac34b4fa65b2a5ba",
  "title": "Standup meeting",
  "note": "bring the release notes",
  "due": "2026-10-23T09:30:00",
  "end": "2026-10-23T10:00:00",
  "repeat": "weekday",
  "lead": 10,
  "done": false,
  "doneAt": null,
  "notified": false,
  "snoozeUntil": null,
  "createdAt": "2026-10-22T20:15:10",
  "completedCount": 3,
  "lastAlert": null
}
```

`due` and `end` are local times in `yyyy-MM-ddTHH:mm:ss`. `end` may be `null` (point-in-time task). `repeat` is one of `none`, `daily`, `weekly`, `monthly`, `weekday`. `lead` is the number of minutes to remind before `due`.

## Commands

| Command | What it does |
| --- | --- |
| `Start-TaskReminder.vbs` | Start silently — this is what the shortcuts use |
| `-Setup` | Create Desktop + Startup shortcuts and enable autostart |
| `-RemoveShortcuts` | Remove both shortcuts (the app itself stays) |
| `-SelfTest` | Run the 28 built-in logic assertions and exit with code 0/1 |
| `-Preview` | Show one sample reminder card, then quit after ~14 seconds |
| `-StartHidden` | Start minimised to the tray (used by the Startup shortcut) |

```powershell
# logic self-test: repeat rules, due checks, snooze, JSON round-trip
powershell -NoProfile -ExecutionPolicy Bypass -File .\TaskReminder.ps1 -SelfTest
```

## Troubleshooting

**Nothing appears after login.** It starts hidden. Look for the blue check icon in the tray (possibly under the `^` overflow); double-click it to open the window.

**The reminder card is covered.** The card is a top-most window, but desktop pets or widgets pinned to the bottom-right corner can sit above it. Move or close the widget.

**Windows or antivirus complains about PowerShell.** The app is an unsigned `.ps1` started with `-ExecutionPolicy Bypass`, which some heuristics flag. The whole program is one readable script — review it, or unblock the file with `Unblock-File .\TaskReminder.ps1`.

**"Already running".** Only one instance is allowed by design; launching again brings the existing window to the front.

**Can I lose my tasks?** Each change is written immediately. Saves retry five times if the file is momentarily locked by OneDrive or an antivirus, and tell you when they fail. A corrupt file is backed up as `tasks.json.broken-<timestamp>` and, if the file is missing, the previous `.old` copy is restored on start-up.

## Known limitations

- The UI is DPI-unaware, so on 150 % display scaling Windows stretches it and it looks slightly soft.
- Monthly repeats advance from the previous date, so a task on the 31st moves to the 28th in February and stays there in later months.
- The interface and tray tooltips are Chinese only.
- No cloud sync or mobile app — by design.

## Changelog

**v1.4** — atomic saves with retries and rollback; fixed snooze interacting with lead time; un-ticking "done" re-arms the reminder; editing or postponing a finished task makes it pending again; the add-box default time no longer freezes at launch; the tray "start with Windows" toggle works; `Space` toggles completion; Windows shutdown is no longer blocked; 28 self-test assertions.

**v1.3** — time-range tasks are judged overdue by their end time; new "in progress" state; column widths fixed.

**v1.2** — unfinished tasks re-alert at a configurable interval; the card shows how long a task has been overdue; new postpone actions.

**v1.1** — real date/time pickers instead of free text (no more silent wrong-time fallback); time ranges; fixed accidental completion when clicking row text.

## License

No license file is included yet. Add one (MIT, Apache-2.0, …) before accepting outside contributions.