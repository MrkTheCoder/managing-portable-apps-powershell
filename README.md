# Managing Portable Apps – PowerShell Script

## Introduction

This script automates the management of portable applications by scanning a designated folder where each portable app is stored in its own subfolder and identified by a `.app` JSON file. It then compares those portable apps against the Start Menu entries on the system, allowing you to review, add or remove shortcuts as needed. The tool gives you a visual UI to manage which portables appear in your Start Menu, making it easier to keep your portable app collection in sync and maintain a clean environment.

## App Window
<img width="860" height="493" alt="image" src="https://github.com/user-attachments/assets/829663cd-9c75-40c9-8e20-02abb0ad757f" />



## Development Process

This project was developed using Windows PowerShell and the .NET Windows Forms libraries (via `System.Windows.Forms` and `System.Drawing`). The workflow included:

* Defining a JSON-based `.app` file specification for each portable application.
* Writing PowerShell functions to scan directories, read JSON, compute file hashes, and query the Windows registry for installed applications.
* Building a user interface (UI) in PowerShell using WinForms: a TreeView, details panel, combo-box filter, and multiple action buttons.
* Iterating the UI layout to use a `TableLayoutPanel` for responsive and consistent placement instead of manual `Location` objects.
* Leveraging UI event handlers (e.g., AfterSelect, AfterCheck) and internal logic for filtering, selection, and action-button enabling.
* Emphasizing separation of concerns by dividing logic into helper functions (e.g., `Update-DetailsTextBox`, `Set-AppNodeIcon`) and application logic.
* Although my PowerShell experience was limited, I guided two AI assistants (ChatGPT + Claude.ai) to generate, refine, and structure the code according to my design, then thoroughly reviewed and modified it to match the intended behaviour and best practices.

## Highlight Features

* 🎯 Tree-based view of portable apps grouped by category
* ✅ Automatic detection of whether an app is installed, has a Start Menu shortcut, or both
* 🔍 Filter options: All / Installed / Portable on StartMenu / Portable not used
* 🧮 Hash-comparison of `.app` file vs actual shortcut `.app` to detect changes
* 🔧 Buttons for mass-select actions (Select All, Unselect All, Invert Selection)
* 📁 Dedicated buttons: Add Shortcut, Remove Shortcut, Create `.app` file, About
* 🧠 Responsive UI layout using a `TableLayoutPanel` (no manual repositioning)

## Detailed Feature Breakdown

| Feature                                                                                | Description                                                                                                                                                                                                           | Benefit                                                                                  |
| -------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| **Tree-based view**                                                                    | Displays each portable app (with grouping via `appGroup`) and uses icons: “portable”, “installed”, “startMenu”                                                                                                        | Provides clear visual status at a glance                                                 |
| **Installation & shortcut detection**                                                  | For each portable wrapper: checks if installed (`IsInstalled`), has a Start Menu shortcut (`HasShortcut`), and if the portable and shortcut `.app` files match via MD5 (`IsBothSame`)                                 | Helps you identify duplicates, outdated shortcuts, or unused portables                   |
| **Filter combobox (“All”, “Installed”, “Portable on StartMenu”, “Portable not used”)** | Allows you to quickly focus on a subset of apps                                                                                                                                                                       | Speeds up management and cleanup                                                         |
| **Detail panel**                                                                       | Shows metadata of the selected app: name, version comparison, group, folder, description                                                                                                                              | Gives context and helps you make informed decisions                                      |
| **Mass-action buttons (Select All / Unselect / Invert)**                               | Lets you easily change the selection state of many nodes at once, with logic to update parent group check states                                                                                                      | Efficient for large collections                                                          |
| **Action buttons (Add Shortcut / Remove Shortcut / Create `.app` / About)**            | Enables targeted operations: adding missing shortcuts, removing unwanted ones, creating wrapper metadata files, and viewing about information. Buttons are dynamically enabled/disabled based on selection and state. | Streamlines workflow and ensures correct operations                                      |
| **Responsive layout with TableLayoutPanel**                                            | All UI elements are arranged in panels with docking and spacing, eliminating manual positioning and resizing code                                                                                                     | More maintainable UI and improved experience on different resolutions                    |
| **Hash comparison for `.app` ↔ shortcut `.app`**                                       | Compares the MD5 of the portable’s `.app` file and the shortcut’s `.app` in Start Menu to detect if they diverge (`IsBothSame`)                                                                                       | Lets you detect when a portable version has changed but the shortcut hasn’t been updated |
| **About dialog**                                                                       | Shows an introduction, author details, and acknowledgement of AI-guided code generation followed by author review                                                                                                     | Transparently credits the development process and gives users context                    |

---

### Technologies & Icons

<img width="64" alt="PowerShell" src="https://upload.wikimedia.org/wikipedia/commons/2/2f/PowerShell_5.0_icon.png" /> <img width="64" alt="WinForm" src="https://upload.wikimedia.org/wikipedia/commons/thumb/0/09/Microsoft_Forms_%282019-present%29.svg/2203px-Microsoft_Forms_%282019-present%29.svg.png" /> <img width="64" alt="JSON" src="https://cdn-icons-png.flaticon.com/512/136/136443.png" /> <img width="64" alt="VSCode" src="https://upload.wikimedia.org/wikipedia/commons/thumb/9/9a/Visual_Studio_Code_1.35_icon.svg/250px-Visual_Studio_Code_1.35_icon.svg.png" /> <img width="64" alt="ChatGPT" src="https://github.com/user-attachments/assets/47bd4ab7-dedd-4576-8054-07aa4832943c" /> <img width="64" alt="Claude.ai" src="https://upload.wikimedia.org/wikipedia/commons/thumb/b/b0/Claude_AI_symbol.svg/250px-Claude_AI_symbol.svg.png" />

* **PowerShell** – core scripting language used for orchestration and UI. ([Icons8][1])
* **Windows Forms (.NET)** – UI framework to build the graphical interface inside PowerShell.
* **JSON** – file format for `.app` wrappers describing each portable application.
* **Visual Studio Code** – the IDE used for writing and editing this PowerShell script.
* **ChatGPT** – one of the AI tools used to help generate and refine the code and logic.
* **Claude AI** – the second AI assistant employed in the development process for analysis and review.


---

## Getting Started

1. Clone the repository:

   ```powershell
   git clone https://github.com/YourUsername/managing-portable-apps-powershell.git
   ```
2. Place your portable applications in subfolders and include `.app` JSON files describing each (with properties like `appName`, `appVersion`, `appStartMenuFolderName`, etc.).
3. Place all `*.ps1` files in the root directory of your main portable apps and execute the script through the evaluated window.
   ```
   [Drive:]PortableAppsFolder/
		├── App1Folder/
		│   ├── SubFolder1/
		│   ├── .app
		│   └── App1.exe
		├── App2Folder/
		│   ├── .app
		│   └── App2.exe
		├── Manage-Portable-Apps.ps1
		└── Create-.app-file.ps1
   ```
   
   ```powershell
   powershell -ExecutionPolicy Bypass -File Manage-Portable-Apps.ps1
   ```
5. Use the UI: filter, select, view details, and perform actions such as adding/removing shortcuts or creating new `.app` wrappers.

---

## Contributing

Contributions, bug reports, and enhancements are welcome. Please fork the project and submit a pull request. Maintain code consistency and update documentation as necessary.

---

## License & Acknowledgements

This script is provided *as-is*. Feel free to adapt or extend it for your portable-apps management workflow. Special thanks to the AI-assistants (code generation: ChatGPT + Claude.ai).
