# VitalSource Desktop Capture

A Python tool to capture pages from the VitalSource Bookshelf desktop application and assemble them into a PDF. This version captures screens directly from the desktop app to avoid online login scripts and CAPTCHA detection.

## Features

- **Relative Clicking**: Automatically adjusts the "Next Page" click location even if you move the Bookshelf window.
- **Global Hotkeys**:
    - **F10**: Emergency Stop (Kill Switch) - stops the process immediately.
    - **F9**: Pause/Resume - toggle the capture loop.
    - **q**: Backup stop key.
- **Smart Resuming**: Automatically detects existing page captures and resumes where it left off.
- **Auto-Cropping**: Automatically removes sidebars and toolbars to keep only the book content.

## Setup

1.  **Install VitalSource Bookshelf** and open your book.
2.  **Install Python** (3.8+ recommended).
3.  **Install Requirements**:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Run the application:
    ```bash
    python vitalsource_desktop.py
    ```
2.  In the GUI:
    - Click **Set Next Button Location**.
    - Hover your mouse over the "Next Page" (>) button in your Bookshelf app.
    - Press the **'n'** key to save the relative location.
3.  Enter the **Total Pages** (or leave blank).
4.  Optionally adjust the **Delay (ms)** between clicks.
5.  Click **Start Capture**.
6.  The tool will:
    - Bring the Bookshelf window to the front.
    - Capture the page.
    - Click "Next".
    - Repeat until done or stopped.
7.  Once finished, the tool will assemble the images into `converted_book.pdf`.

## Disclaimer

This software is for personal use and backup purposes only. Please respect copyright laws and VitalSource terms of service. Only use this for books you have legally purchased.

## Troubleshooting

- **Window Not Found**: Make sure you have the **VitalSource Bookshelf Desktop Application** installed and open. This tool does **NOT** work with the web browser reader.
- **Offsets are wrong**: If the mouse clicks the wrong place, try re-setting the "Next Button" location. Ensure your display scaling (DPI) is consistent or set to 100% if you have issues.
- **Screenshots are black**: Ensure the Bookshelf app is not minimized when capturing.
