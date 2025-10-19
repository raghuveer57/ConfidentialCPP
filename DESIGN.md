# Design Document — ConfidentialCPP

Last updated: 2025-10-19

Purpose: capture the high-level design, architecture, module responsibilities, data flows, build and run requirements, security considerations, and suggested next steps for the ConfidentialCPP project (native Win32 C++ application).

## Overview

ConfidentialCPP is a native Windows application written in C++ using the Win32 API and Visual C++ toolchain. Its primary runtime behavior is:

- Monitor top-level windows and identify target application windows (Word, Excel, PowerPoint, Acrobat, etc.).
- Attach and show a small floating UI (button) adjacent to target windows. When interacted with, the UI exposes a pane that queries or toggles a document confidentiality flag by invoking an external helper executable (`WordDocumentUpdater.exe`).
- Provide a tray icon and minimal context-menu control for the application lifecycle.

The code is contained in the `Confidential/` project and is built through `Confidential.sln`. An installer project exists under `Installer/`.

## High-level architecture

- Single-process, single Windows GUI application.
- Main message loop manages timers and enumerates windows periodically (uses `EnumWindows` and `GetWindowThreadProcessId`).
- Per-target-window UI composed of two window classes:
  - `ButtonWindowClass` — owner-drawn small floating button that shows/hides and handles click state.
  - `PaneWindowClass` — a popup pane shown when the button is toggled; contains controls and initiates actions.
- Communication with external logic is implemented by spawning subprocesses (via `CreateProcess`) to run `WordDocumentUpdater.exe` with `get` and `toggle` commands.
- WMI (COM) calls are used to obtain the original process command-line and derive file paths.

## Modules and responsibilities

- main (wWinMain / WndProc)
  - App initialization, class registration, timer setup, message loop.
  - High-frequency timer (10 ms as currently configured) triggers `UpdateButtonPosition()`.

- Window Enumerator (EnumWindowsProc, UpdateButtonPosition)
  - Scans top-level windows, extracts process names and class names, and determines whether to create or update floating UI elements.

- UI Layer (ButtonWndProc, PaneWndProc)
  - `ButtonWindowClass` draws custom rounded button icon and handles click/toggle state in `g_checkMap`.
  - `PaneWindowClass` creates a pane with controls (buttons, static text) and periodically queries the external helper for state.

- Process/OS Utilities
  - `GetProcessName`, `GetCommandLineFromProcess`, `GetExecutablePath/Directory`, `RunCommand`, `RunPowerShellCommand`.
  - File read/write helpers using Win32 handles.

- External Helper
  - `WordDocumentUpdater.exe` (not included in repository root) — executed with `get` or `toggle` to query or change the per-document confidentiality flag. The main application depends on this helper for its core document operations.

## Data flows

- Detection flow:
  - `EnumWindows` -> for each window call `GetWindowThreadProcessId` -> `GetProcessName` -> if in `g_targetProcesses` create/update button/pane

- UI action flow:
  - User clicks floating button -> set check state and show `PaneWindowClass`
  - Pane triggers `WordDocumentUpdater.exe get` to observe current document state
  - Pane displays the state and, if the user clicks action, runs `WordDocumentUpdater.exe toggle` to change the state
  - Pane updates UI based on helper's stdout (expects leading '0' or '1')

- Command-line / path extraction flow:
  - When creating a pane, code calls `GetCommandLineFromProcess(processId)` using WMI to derive the document path from the process command-line. This value is stored in `g_pathMap[pane]`.

## Important design decisions & constraints

- Polling frequency: currently a timer with 10 ms interval (very frequent) updates positions. This may impact CPU. Consider switching to a longer interval (100–500 ms) or subscribing to window events (SetWinEventHook) for better efficiency.

- UI implementation: owner-drawn windows with manual double-buffering for rendering. This keeps the dependency footprint low but increases code complexity. Consider migrating to a light UI framework (e.g., WinUI, WTL, or a small C++ GUI helper) if more features are required.

- External helper usage: The app shells out to `WordDocumentUpdater.exe` for document operations. This keeps the main process simple and reduces third-party dependency exposure. However, it requires that the helper executable is available and trustworthy. Consider in-process library integration if tighter control or performance is desired.

- WMI usage: WMI is used to get the command-line for a process. WMI requires COM initialization and can be slow. Consider using QueryFullProcessImageName or other techniques for process introspection where available.

## Build & run

Prerequisites:
- Windows 10 or later
- Visual Studio 2019/2022 with Desktop C++ workload

Build steps:
1. Open `Confidential.sln` in Visual Studio.
2. Select platform (x64 or Win32) and configuration (Debug/Release).
3. Build solution (Ctrl+Shift+B).

Run:
- Run `Confidential.exe` from Visual Studio or directly from the output folder.
- Ensure `WordDocumentUpdater.exe` is present next to the EXE or adjust the code that discovers it (`GetExecutableDirectory`).

## Testing

- There are no automated tests in the repo. Suggested test plan:
  - Unit tests for utility functions (path extraction, string conversions) using a small C++ test framework.
  - Integration test: run the app and simulate target applications (Word/Excel) with mock processes to verify creation/hide/show behavior.
  - Manual testing: verify that the floating UI appears for target windows and that `get` / `toggle` commands are executed correctly and parsed.

## Security considerations

- Running external executables should be done carefully; ensure `WordDocumentUpdater.exe` is signed/trusted.
- Using `CreateProcess` with unescaped command lines or relying on current working directory may be vulnerable to path injection if untrusted data is used. Currently the code builds a quoted path which mitigates common cases but review argument escaping thoroughly.
- WMI usage and COM initialization needs proper error-handling and cleanup (CoUninitialize). Avoid leaking COM pointers.
- File read/write uses full read/write access to the file; if this runs with elevated privileges it could modify files unexpectedly. Ensure least-privilege operation and validate file paths before writing.

## Performance & resource usage

- High-frequency polling (10 ms) is a likely cause of unnecessary CPU usage. Increase the timer interval or switch to event-driven updates.
- Spawning helper processes for every UI check may be expensive. Consider caching results for short periods or using a single long-running helper process it communicates with via IPC (named pipes) for lower overhead.

## Known issues & TODOs

- [ ] Replace 10 ms timer with a more reasonable interval or event-driven approach (SetWinEventHook)
- [ ] Add robust error handling around COM/WMI calls and process creation
- [ ] Add unit tests for utilities and integration tests for UI flow
- [ ] Consider in-process integration with document handling logic (or provide a signed, versioned helper)
- [ ] Improve security by validating paths and signing the helper executable
- [ ] Add logging and an option for verbose diagnostic output
- [ ] Add installer migration plan (migrate `.vdproj` to WiX) and CI build steps

## Implementation notes / important functions

- `wWinMain` — application entry and initialization
- `UpdateButtonPosition` — called on timer, enumerates windows and triggers UI creation
- `CreateOrUpdateButton` — creates and positions `ButtonWindowClass` and `PaneWindowClass` instances
- `GetCommandLineFromProcess` — uses WMI to query Win32_Process.CommandLine
- `RunCommand` / `RunPowerShellCommand` — spawn child process and read stdout

## Diagrams (text)

Window/Process Interaction:

[EnumWindows] -> For each window -> GetProcessName -> If in target set -> Create/Update button & pane

UI Interaction:

[User click button] -> Set check state -> Show pane -> Pane runs `WordDocumentUpdater.exe get` -> Display state -> [User action] -> Pane runs `WordDocumentUpdater.exe toggle`

## Next steps I can take for you

- Produce a `DESIGN.md` PR with a targeted set of changes to address the TODOs (for example: change timer interval & add SetWinEventHook fallback, add logging, add basic unit tests).
- Create a CI workflow (GitHub Actions) that builds the solution on Windows and runs optional unit tests.
- Create a small test harness or mock `WordDocumentUpdater.exe` to allow safe local testing without touching real Office documents.

If you want me to continue, tell me which next step to take and I'll implement it.