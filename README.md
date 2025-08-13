# ✨ WinPnpUtilGui - GUI for PnPUtil on Windows

A simple C# Windows Forms application providing a GUI for managing device drivers via `pnputil.exe`.

## Features 🚀

*   **Enumerate Drivers:** List all installed third-party drivers.
*   **Sort Drivers:** Click column headers to sort (toggle ascending/descending).
*   **Search/Filter:** Filter by Provider (case-insensitive, fuzzy) using the search box.
*   **Uninstall Drivers:** Right-click selected drivers to uninstall (supports multi-select).

## Usage 💻

### Requirements
*   Windows OS (7+), .NET Framework 4.7.2+.
*   **Run as Administrator** for all functions.

### Quick Start
1.  Run `PnpUtilGui.exe` **as Administrator**.
2.  Click **"Enumerate Drivers"** to populate the list (defaults to Provider sort).
3.  **Uninstall:** Select drivers, right-click, choose "Uninstall Driver", confirm.

## Building (for Developers) 🛠️
1.  Open `PnpUtilGui.sln` in Visual Studio.
2.  Set "Platform target" to "x64" or "Any CPU".
3.  Build the solution.

## Troubleshooting 💡
*   **"File not found" / No drivers:** Ensure "x64" build and **Run as Administrator**.
