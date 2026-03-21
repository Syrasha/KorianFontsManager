# KorianFontsManager

A simple font manager for Windows built with Python and Tkinter.

## How to Build the Executable

To create a standalone `.exe` that you can share with others, use the provided PowerShell script:

1.  Open a PowerShell terminal in this directory.
2.  Run the following command:
    ```powershell
    .\build_exe.ps1
    ```
    *If you get an execution policy error, run:*
    ```powershell
    PowerShell -ExecutionPolicy Bypass -File .\build_exe.ps1
    ```

The script will:
-   Check for a virtual environment (`.venv`).
-   Install required dependencies (`Pillow`, `PyInstaller`).
-   Generate a single-file executable in the `dist` folder named `KorianFontsManager.exe`.

## Dependencies
-   Python 3.x
-   Pillow (for image/font handling)
-   PyInstaller (for packaging)

## Usage Note
The application uses `config.json` to store your favorites and projects. When you run the `.exe` for the first time, it will create this file in the same directory if it doesn't already exist.
