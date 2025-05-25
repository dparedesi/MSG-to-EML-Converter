# run_app_launcher.py
import subprocess
import os
import sys
import webbrowser
import time # To give the server a moment to start

def get_path(filename):
    if hasattr(sys, "_MEIPASS"): # PyInstaller temp folder
        return os.path.join(sys._MEIPASS, filename)
    return filename # Running as script

if __name__ == '__main__':
    app_py_path = get_path('app.py') 
    streamlit_cmd = "streamlit" # Assume streamlit is in PATH or bundled correctly
    
    # Define the port (make sure it matches what Streamlit uses or what you configure)
    port = "8501" 
    url = f"http://localhost:{port}"

    # Command to run Streamlit
    # Using --server.headless=true can be good as it implies Streamlit shouldn't try to open a browser itself.
    # We are handling the browser opening.
    # Using --server.runOnSave=false can also be good for a "production" feel of the local app.
    cmd = [
        streamlit_cmd, "run", app_py_path, 
        "--server.headless=true", 
        "--server.port", port,
        "--server.runOnSave=false" 
    ]

    print(f"Attempting to launch Streamlit with command: {' '.join(cmd)}")
    print(f"If the browser doesn't open automatically, please open it to: {url}")
    print(f"This window shows server logs. Close this window to stop the application.")
    
    try:
        # Start Streamlit as a background process
        process = subprocess.Popen(cmd)
        
        # Give the server a few seconds to start up before trying to open the browser
        print("Waiting for Streamlit server to start...")
        time.sleep(3) # Adjust as needed; 2-5 seconds is usually enough

        # Check if the server seems to be running (optional, but good practice)
        # A more robust check would be to try making an HTTP request to the health endpoint
        # For now, we'll assume it starts within the sleep period.

        print(f"Opening browser to {url}...")
        webbrowser.open(url) # Open the URL in the default web browser
        
        process.wait() # Wait for the Streamlit process to terminate (e.g., user closes the terminal)

    except FileNotFoundError:
        print(f"Error: Could not find '{streamlit_cmd}'. Make sure Streamlit is installed and in your PATH, or that PyInstaller bundled it correctly.")
        print("If running the bundled app, this indicates a packaging issue.")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"An error occurred: {e}")
        input("Press Enter to exit...")