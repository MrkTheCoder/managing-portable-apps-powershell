# Managing Portable Apps – PowerShell Script

***

## Introduction

Managing portable applications can quickly become a headache—especially when dealing with shortcuts. A typical workflow looks like this: you gather your portable apps in a folder, manually create shortcuts, copy them into the Start Menu, and hope everything keeps working. But the moment you move that folder, switch drives, or copy the setup to another PC, all your shortcuts break. Suddenly, you’re fixing paths one by one, recreating shortcuts, or digging through folders to launch apps directly.

This project solves that problem by introducing a smarter, dynamic system for managing portable applications.

At the core of the system is a **`.app` JSON package file**, placed inside each portable application's folder. This file defines all metadata and Start Menu shortcuts for the app using dynamic path variables—most importantly **`[.app_path]`**, which always resolves to the portable application's real location at runtime. This means your shortcuts **never break**, even if you move the folder to another computer, drive, or directory.

The main script scans your portable apps directory, loads each `.app` package, and allows you to:

* **Create Start Menu shortcuts dynamically**
* **Add or remove shortcuts with a visual UI**
* **Detect existing installed applications**
* **Compare Start Menu entries with your portable apps**
* **Keep your system clean, consistent, and fully portable**

To make things even easier, the project includes a dedicated tool—

### **Shortcut-to-.app Converter**

This helper script takes your **manually created .lnk shortcuts** and converts them into a reusable dynamic `.app` package file. Instead of writing JSON by hand, you simply create shortcuts once, run the converter, and the tool generates a portable `.app` definition that works anywhere. This makes it effortless to migrate apps, replicate configurations, or build `.app` packages for all your portable tools.

Whether you're maintaining a large portable apps collection or just want reliable Start Menu shortcuts, this project provides a seamless, flexible, and time-saving solution for everyone.


***

## App Windows

<p align="center">
	<img width="800" alt="gif" src="https://github.com/user-attachments/assets/53873601-5d8d-4de0-8d9f-571a238109a6?raw=true" />
</p>
<p align="center">👆 animated GIF of the main app window demonstration 👆</p>
<br>
<br>
<br>
<p align="center">
	<img width="686" height="553" alt="image" src="https://github.com/user-attachments/assets/4881e8ad-e3ea-4331-b386-e54bff8266b0" />
</p>

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
