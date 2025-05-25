# MSGtoEMLConverter.spec

# This is a simplified example; Streamlit packaging can be complex.
# You might need to add specific hooks for Streamlit or its dependencies.

block_cipher = None

a = Analysis(
    ['run_app_launcher.py'], # We package the launcher
    pathex=[], # Add your project directory if needed: ['/path/to/your/project']
    binaries=[],
    datas=[
        ('app.py', '.'), # Include app.py at the root of the packaged app
        ('msg_converter_core.py', '.') # Include your core logic
        # Add other data files if any (e.g., images, templates)
    ],
    hiddenimports=[
        'streamlit.web.server.server', # Common hidden import for Streamlit
        'altair', 'pandas', 'pyarrow', # Common Streamlit dependencies, add if used
        'watchdog', # Often needed by Streamlit
        # Add other hidden imports your app or extract-msg might need
        'extract_msg', 
        'extract_msg.utils', # Be explicit with submodules if issues arise
        'extract_msg.msg_classes',
        'email.mime', # Explicitly include email submodules if needed
        'email.header',
        'email.encoders',
        'email.utils',
        # ... other email submodules if issues
        'chardet', 'olefile', 'compressed_rtf', # Dependencies of extract-msg
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Collect Streamlit's own static files (important!)
# This path might vary based on your Python environment.
# Find where Streamlit is installed: `pip show streamlit` -> Location
# Then navigate to `streamlit/static`
# Example (adjust this path!):
# st_static_path = 'path/to/your/python/Lib/site-packages/streamlit/static' 
# if os.path.exists(st_static_path):
#    a.datas += Tree(st_static_path, prefix='streamlit/static')
# else:
#    print(f"WARNING: Streamlit static path not found at {st_static_path}")

# A more reliable way using importlib.resources or pkg_resources if possible,
# or a PyInstaller hook for Streamlit is best.
# For now, let's hope PyInstaller's default hooks handle Streamlit's static files.
# If UI elements are missing in the packaged app, this is the first place to look.

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MSGtoEMLConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True, # Compresses the executable, can be slow
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True, # True for the launcher to show Streamlit server logs
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='MSGtoEMLConverter.ico' # Add an icon if you have one
)