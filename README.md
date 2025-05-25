# MSG to EML Converter ✉️✨

A powerful local utility to convert Microsoft Outlook `.msg` files into standard `.eml` files. This tool intelligently handles nested `.msg` attachments, converting them and embedding them as `message/rfc822` parts within the main EML. This means you can view the entire email chain, including original attachments, in most modern email clients that support this standard.

The application provides a user-friendly interface built with Streamlit and can be packaged into a standalone executable for easy distribution.

## Features

*   **MSG to EML Conversion:** Converts the main `.msg` file to a fully compliant `.eml` file.
*   **Nested MSG Handling:** Recursively processes `.msg` files attached within other `.msg` files.
*   **Embedded EMLs:** Converts nested `.msg` attachments into `.eml` format and embeds them as `message/rfc822` MIME parts in the parent EML. This allows them to be viewed inline or as attached EMLs in compatible email clients.
*   **Preserves Standard Attachments:** Regular file attachments (non-MSG) are preserved in the final EML.
*   **User-Friendly Interface:** A simple web-based UI powered by Streamlit for easy file uploading and conversion.
*   **Local & Offline:** Runs entirely on your local machine. No internet connection or external servers required after setup/packaging.
*   **Standalone Executable:** Can be packaged into a single executable for colleagues who don't have Python installed.

## How to Use (Packaged Application)

This is the recommended way for most users.

1.  **Download:** Obtain the `MSGtoEMLConverter` executable (e.g., `MSGtoEMLConverter.exe` on Windows or `MSGtoEMLConverter` on macOS/Linux) from the `dist` folder.
2.  **Run:** Double-click the executable.
    *   A terminal window will open (this shows server logs; you can usually minimize it).
    *   Your default web browser should automatically open to the application's interface (usually `http://localhost:8501`). If it doesn't, manually open your browser and navigate to that URL.
3.  **Upload:** Use the file uploader in the web interface to select the `.msg` file you want to convert.
4.  **Convert:** Click the "Convert" button.
5.  **Download:** Once the conversion is complete, a download button will appear for your new `.eml` file. The logs on the page will show the conversion progress.
6.  **Close:** To stop the application, simply close the terminal window that opened when you launched the executable.

## How to Run (From Source - For Developers)

If you have Python and want to run the application directly from the source code:

1.  **Prerequisites:**
    *   Python 3.8 or higher.
    *   `pip` (Python package installer).

2.  **Clone the Repository (if applicable) or Download Files:**
    Ensure you have `app.py`, `msg_converter_core.py`, and `run_app_launcher.py`.

3.  **Create a Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

4.  **Install Dependencies:**
    If a `requirements.txt` file is provided:
    ```bash
    pip install -r requirements.txt
    ```
    Otherwise, install manually:
    ```bash
    pip install streamlit extract-msg
    # (Ensure extract-msg version matches what was used, e.g., extract-msg==0.41.2)
    ```

5.  **Run the Streamlit Application:**
    Navigate to the project directory in your terminal and run:
    ```bash
    streamlit run app.py
    ```
    This will start the Streamlit development server, and your browser should open to the application.

## Building the Executable (For Developers)

To package the application into a standalone executable using PyInstaller:

1.  **Install PyInstaller:**
    ```bash
    pip install pyinstaller
    ```

2.  **Navigate to Project Directory:**
    Open your terminal in the root directory of the project (where `run_app_launcher.py`, `app.py`, and `msg_converter_core.py` are located).

3.  **Create/Update the `.spec` File:**
    A `MSGtoEMLConverter.spec` file is used to configure the PyInstaller build. If one isn't provided, you can generate a basic one and then modify it:
    ```bash
    pyi-makespec --name MSGtoEMLConverter run_app_launcher.py
    ```
    Then, edit `MSGtoEMLConverter.spec` to include necessary `datas` (like `app.py`, `msg_converter_core.py`), `hiddenimports` (for Streamlit, extract-msg, and their dependencies), and potentially hooks for Streamlit if UI elements are missing. Refer to the comments within a provided `.spec` file for guidance.

    **Example `hiddenimports` that might be needed:**
    ```python
    hiddenimports=[
        'streamlit.web.server.server', 'altair', 'pandas', 'pyarrow', 'watchdog',
        'extract_msg', 'extract_msg.utils', 'extract_msg.msg_classes',
        'email.mime', 'email.header', 'email.encoders', 'email.utils',
        'chardet', 'olefile', 'compressed_rtf',
        # Add any other modules your application or its dependencies might use implicitly.
    ]
    ```

4.  **Build the Executable:**
    ```bash
    pyinstaller MSGtoEMLConverter.spec
    ```
    This command will create a `build` folder and a `dist` folder. The executable will be inside `dist/MSGtoEMLConverter`.

5.  **Test:** Run the executable from the `dist` folder.

## Troubleshooting Packaged App

*   **UI Looks Plain/Broken:** This often means Streamlit's static assets (CSS, JavaScript) were not bundled correctly. You may need to:
    *   Ensure PyInstaller hooks for Streamlit are active (e.g., by installing `pyinstaller-hooks-contrib` if it has a Streamlit hook).
    *   Manually add Streamlit's `static` directory to the `datas` section in your `.spec` file using `Tree()`. Example: `a.datas += Tree('/path/to/streamlit/static', prefix='streamlit/static')` (adjust the source path).
*   **`ModuleNotFoundError` on Launch:** A required module was not included. Add it to the `hiddenimports` list in the `.spec` file and rebuild.
*   **`FileNotFoundError: streamlit` (from launcher):** PyInstaller might not have found or correctly bundled the `streamlit` command-line entry point. Ensure `streamlit` is installed in the environment PyInstaller is using. Sometimes, specifying the full path to the bundled Python interpreter for the `streamlit` command can help: `cmd = [sys.executable, "-m", "streamlit", "run", ...]`.
*   **Large Executable Size:** PyInstaller bundles a Python interpreter and dependencies. Using UPX (enabled by `upx=True` in the `.spec` file) can help compress the executable, but it can also make the build process slower or occasionally cause issues with antivirus software.

## Dependencies

*   **Python 3.8+**
*   **extract-msg:** For parsing `.msg` files.
*   **Streamlit:** For the web-based user interface.

(See `requirements.txt` if provided for specific versions).

