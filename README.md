# Managing Portable Apps – PowerShell Script

***

## Introduction

This script automates the management of portable applications by scanning a designated folder where each portable app is stored in its own subfolder and identified by a **`.app` JSON file**. The script allows **users to create this marker file** for new applications. It then compares those portable apps against the Start Menu entries on the system, allowing you to review, add, or remove shortcuts as needed. Because the shortcuts are based on the `.app` file's dynamic pathing logic, they **adapt even if you move the portable application folder** to a new location. The tool gives you a visual UI to manage which portables appear in your Start Menu, making it easier to keep your portable app collection in sync and maintain a clean environment.

***

## App Window

<p align="center">
	<img width="836" height="493" alt="image" src="https://github.com/user-attachments/assets/2468c691-afb5-4df1-82f8-a48956638a2e" />
</p>

***

## Highlight Features

* **JSON-Based Specification:** Uses a standard **`.app` JSON file** as a specification marker to define each portable application and its shortcut metadata.
* **Dynamic Path Resolution:** All generated shortcuts are dynamic, ensuring they remain valid and functional even if the portable application folder is **moved to a new drive or system**.
* **Intelligent Syncing:** The script scans your local portable apps and checks the Windows Start Menu and installed applications, providing **visual status indicators** (Installed, Shortcut OK, Shortcut Modified, Missing).
* **Full UI Management:** Applications are managed through a comprehensive **WinForms GUI** featuring a TreeView, filtering controls, a details panel, and immediate action buttons.
* **Core Actions:** Supports **adding/updating shortcuts**, **removing obsolete shortcuts**, and **generating new `.app` definition files** for newly added portable apps.
* **Clean Architecture:** Code leverages PowerShell functions for **directory scanning**, **JSON handling**, and **UI event management**, promoting separation of concerns and maintainability.

***

## `.app` JSON File Structure

Each portable application is identified by a JSON-based **`.app` marker file** located in its root folder. This file contains metadata and definitions for one or more Start Menu shortcuts.

The key feature is the use of the dynamic variable `[.app_path]` within the file. This placeholder is automatically resolved by the script to the current, absolute path of the portable application's root folder at runtime, enabling flexible and efficient shortcut creation regardless of the app's location.

Here is a sample file demonstrating the structure and use of `[.app_path]`:

```json
{
  "appName": "Screen Copy",
  "appVersion": "3.3.1",
  "appGroup": "Mobile Suite",
  "appDescription": "scrcpy is a free and open-source app that allows screen mirroring and control of \nAndroid devices via USB or Wi-Fi. It supports keyboard and mouse interaction.",
  "appInstallRegistryData": "",
  "appStartMenuFolderName": "Screen Copy (Mobile Remote Control)",
  "shortcuts": [
    {
      "name": "[Open DOS] scrcpy x64 (Portable)",
      "target": "%SystemRoot%\\System32\\cmd.exe",
      "arguments": "[.app_path]\\",
      "workingDirectory": "[.app_path]\\",
      "icon": ",0",
      "windowStyle": 1,
      "description": ""
    },
    {
      "name": "scrcpy x64 (Portable)",
      "target": "[.app_path]\\scrcpy-win64\\scrcpy.exe",
      "arguments": "",
      "workingDirectory": "[.app_path]\\scrcpy-win64",
      "icon": ",0",
      "windowStyle": 1,
      "description": ""
    },
    {
      "name": "[Open DOS] scrcpy x86 (Portable)",
      "target": "%SystemRoot%\\System32\\cmd.exe",
      "arguments": "[.app_path]\\",
      "workingDirectory": "[.app_path]\\",
      "icon": ",0",
      "windowStyle": 1,
      "description": ""
    },
    {
      "name": "scrcpy x86 (Portable)",
      "target": "[.app_path]\\scrcpy-win32\\scrcpy.exe",
      "arguments": "",
      "workingDirectory": "[.app_path]\\scrcpy-win32",
      "icon": ",0",
      "windowStyle": 1,
      "description": ""
    }
  ]
}
```

***

## Development Process

This project was developed using **Windows PowerShell** and the **.NET Windows Forms libraries** (via `System.Windows.Forms` and `System.Drawing`). Although my PowerShell experience was limited, I guided two AI assistants (**ChatGPT** and **Claude.ai**) to generate, refine, and structure the code according to my design, then thoroughly reviewed and modified it to ensure it matched the intended behavior and best practices. Key architectural decisions included implementing the UI using a `TableLayoutPanel` for responsive layout and emphasizing the **separation of concerns** within the PowerShell functions.

***

## Technologies & Development

[![PowerShell](https://img.shields.io/badge/PowerShell-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://docs.microsoft.com/en-us/powershell/)
[![Windows Forms](https://img.shields.io/badge/.NET%20Windows%20Forms-512BD4?style=for-the-badge&logo=.net)](https://docs.microsoft.com/en-us/dotnet/desktop/winforms/?view=netdesktop-6.0)
[![JSON](https://img.shields.io/badge/Data%20Format-JSON-000000?style=for-the-badge&logo=json&logoColor=white)](https://www.json.org/)
[![VS Code](https://img.shields.io/badge/Editor-VS%20Code-007ACC?style=for-the-badge&logo=visual-studio-code&logoColor=white)](https://code.visualstudio.com/) 
[![ChatGPT Assisted](https://img.shields.io/badge/Code%20Assisted%20By-ChatGPT-6AA299?style=for-the-badge&logo=openai&logoColor=white)](https://openai.com/chatgpt)
[![Claude AI Assisted](https://img.shields.io/badge/Code%20Assisted%20By-Claude%20AI-5437C0?style=for-the-badge)](https://www.claude.ai/)

***
## Getting Started

1.  Download ZIP or Clone the repository:

    ```powershell
    git clone https://github.com/MrkTheCoder/managing-portable-apps-powershell.git
    ```
2.  Place your portable applications in subfolders and ensure each includes a **`.app` JSON file** describing the application (use the included `Create-.app-file.ps1` script to generate these).

3.  Ensure all `*.ps*` files are in the root directory of your main portable apps folder. The expected structure is:

    ```
    [Drive:]PortableAppsFolder/
    ├── App1Folder/
    │   ├── SubFolder1/
    │   ├── .app
    │   └── App1.exe
    ├── App2Folder/
    │   ├── .app
    │   └── App2.exe
    ├── Manage-Portable-Apps.ps1
    ├── Manage-Portable-Apps.UI.psm1
    └── Create-.app-file.ps1
    ```

4.  Execute the main script from the PowerShell window:

    ```powershell
    powershell -ExecutionPolicy Bypass -File Manage-Portable-Apps.ps1
    ```
5.  Use the UI to filter, select, view details, and perform actions such as adding/removing shortcuts or creating new `.app` wrappers.

***

## Contributing

Contributions, bug reports, and enhancements are welcome. Please fork the project and submit a pull request. Maintain code consistency and update documentation as necessary.

***

## License & Acknowledgements

This script is provided *as-is*. Feel free to adapt or extend it for your portable apps management workflow. Special thanks to the **AI-assistants** (code generation: ChatGPT + Claude.ai) for their instrumental role in the development process.
