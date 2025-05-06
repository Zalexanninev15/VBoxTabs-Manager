# VBoxTabs Manager

[![](https://img.shields.io/badge/platform-Windows-informational)](https://github.com/Zalexanninev15/VBoxTabs-Manager)
[![](https://img.shields.io/badge/PySide6-6.9.0-6F56AE?logo=qt)](https://pypi.org/project/PySide6/)
[![](https://img.shields.io/badge/written_on-Python-%233776AB.svg?logo=python)](https://www.python.org/)
[![](https://img.shields.io/github/v/release/Zalexanninev15/VBoxTabs-Manager)](https://github.com/Zalexanninev15/VBoxTabs-Manager/releases/latest)
[![](https://img.shields.io/github/downloads/Zalexanninev15/VBoxTabs-Manager/total.svg)](https://github.com/Zalexanninev15/VBoxTabs-Manager/releases)
[![](https://img.shields.io/github/last-commit/Zalexanninev15/VBoxTabs-Manager)](https://github.com/Zalexanninev15/VBoxTabs-Manager/commits/main)
[![](https://img.shields.io/github/stars/Zalexanninev15/VBoxTabs-Manager.svg)](https://github.com/Zalexanninev15/VBoxTabs-Manager/stargazers)
[![](https://img.shields.io/github/forks/Zalexanninev15/VBoxTabs-Manager.svg)](https://github.com/Zalexanninev15/VBoxTabs-Manager/network/members)
[![](https://img.shields.io/github/issues/Zalexanninev15/VBoxTabs-Manager.svg)](https://github.com/Zalexanninev15/VBoxTabs-Manager/issues?q=is%3Aopen+is%3Aissue)
[![](https://img.shields.io/github/issues-closed/Zalexanninev15/VBoxTabs-Manager.svg)](https://github.com/Zalexanninev15/VBoxTabs-Manager/issues?q=is%3Aissue+is%3Aclosed)
[![](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![](https://img.shields.io/badge/Donate-FFDD00.svg?logo=buymeacoffee&logoColor=black)](https://z15.neocities.org/donate)

## Screenshot

![VBoxTabs Manager Screenshot](./Screenshot.png)

## Description

**VBoxTabs Manager** is an application that allows you to collect all running VirtualBox virtual machines into one tabbed window (window grouping). All elements and functionality of VirtualBox windows are fully preserved.

## Features

-   **Automatic Detection**: Finds running VirtualBox virtual machines automatically.
-   **Tabbed Interface**: Combines VM windows into tabs with easy switching.
-   **Full VM Functionality**: Preserves the complete functionality of the original VirtualBox window within the tab.
-   **Tab Management**:
    -   Rename tabs for better identification.
    -   Detach VMs back into separate windows using the toolbar or context menu.
    -   Reorder tabs using drag and drop.
    -   Forcefully close a VM window using the toolbar, context menu, or middle mouse click on the tab.
-   **Toolbar Actions**: Quick access buttons for:
    -   Refreshing the VM list.
    -   Attaching all available VMs.
    -   Detaching the current VM.
    -   Force closing the current VM window.
    -   Force closing all available VMs.
    -   Renaming the current tab.
    -   Opening the main VirtualBox application.
-   **Settings**: The program can be conveniently customized with the help of the settings window.
-   **Context Menu**: Right-click a tab for quick actions (rename, detach, close).
-   **Multilingual Title Support**: Recognizes both English "[Running]" and Russian "[Работает]" in window titles.
-   **Clean Exit**: Automatically detaches all VMs when the application is closed.

## Requirements

- OS: Windows 10/11
- VirtualBox 7.0 or higher (It also works with earlier versions 5 and 6)
- No dependencies required for the standalone build (see below for running from source)

## Roadmap (I will be very glad of your help)

- [ ] Linux support.
- [ ] View some information about running virtual machines. The information will be collected directly from the VirtualBox console implementation, most likely it will be.
- [ ] Support for multiple windows in one tab with different layout.
- [ ] Viewing a preview of the selected window when hovering over the tab.
- [ ] Quickly deattach windows and attach with a simple drag and drop. I don't know how to implement it correctly yet, as there may be problems of accidental attachment and deattachment.

## Installation and Usage

### 1. Using the Release (Recommended)

- Download the latest ready-to-use EXE in archive (because the dependency files are already archived) from the [releases page](https://github.com/Zalexanninev15/VBoxTabs-Manager/releases/latest).
- Run `VBoxTabsManager.exe` — **no installation and no dependencies are required in the system**.

### 2. Running from Source

1. Install Python 3.13+ (maybe, 3.8+).
2. Install dependencies:

    ```shell
    pip install -r requirements.txt
    ```

    🌈 [test] If you want more themes available, you can additionally install *qt-themes*:

    ```shell
    pip install qt-themes
    ```

3. Download or clone this repository.

    ```shell
    git clone https://github.com/Zalexanninev15/VBoxTabs-Manager
    ```

4. Start the required virtual machines in VirtualBox.
5. Run the application:

    ```batch
    python ./VBoxTabs-Manager.py
    ```

All detected running virtual machines will be added to the tabs.

### Interesting stuff

-   **Switch between VMs**: Click the corresponding tab. You can move and rename the tabs however you like.
-   **Detach a VM**: Click the "Detach current VM" button in the toolbar, or right-click the tab and select "Detach".
-   **Force Close VM Window**: Select the tab, click the "Close current VM window" button, middle-click the tab, or right-click the tab and select "Close window". *Warning: This terminates the VM process without graceful shutdown.*
-   **Force Close all available VMs** with process `VBoxSVC`, but not with service.
-   **Open VirtualBox Manager**: Click the "Open VirtualBox main application" button in the toolbar. The path to the executable file is specified in the settings, but the default path is set to the default path from the VirtualBox installer.
-   **Context Menu**: Right-click any tab for quick actions (Rename, Detach, Close Window). Hold **Ctrl** while right-clicking to open the menu for an inactive tab without switching to it.
-  **Automatic Refresh**: Periodically checks for new or closed VM windows. This parameter is adjustable in the settings.
-  **Themes**: You can change the design themes in the settings.

## Troubleshooting

- **Virtual machine is not detected**: make sure the virtual machine is running and has "[Running]" or "[Работает]" in the title.
- **Problems displaying window content**: try resizing the main application window.
This often happens if the windows of virtual machines change themselves, which often happens when starting the system in a virtual machine and setting the resolution.

## Building EXE (with files)

A GitHub Action is provided to automatically build a Windows x64 executable file with other files (no dependencies required) on every commit that changes `VBoxTabs-Manager.py`. The resulting build is published as a release with the version taken from the application.

Use the `build.cmd` script to create executable file with other files.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
