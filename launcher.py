# launcher.py - Streamlit in-process launcher
import sys
import os
import webbrowser
import threading

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def open_browser():
    """Open browser after a short delay"""
    import time
    time.sleep(3)
    try:
        webbrowser.open("http://localhost:8501")
        print("✓ Browser opened at http://localhost:8501")
    except Exception as e:
        print(f"✗ Could not open browser: {e}")

def run_streamlit(app_path):
    """Run Streamlit app in-process to avoid relaunch loops"""
    try:
        from streamlit.web import cli as stcli
    except ImportError:
        print("✗ ERROR: Streamlit not found in current environment")
        input("Press Enter to exit...")
        return

    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_PORT"] = "8501"
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "localhost"
    os.environ["STREAMLIT_BROWSER_GATHERUSAGESTATS"] = "false"

    sys.argv = [
        "streamlit", "run", app_path,
        "--server.port=8501",
        "--server.address=localhost",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ]

    print("\n" + "=" * 60)
    print("Starting Streamlit server...")
    print("The app will open in your browser shortly.")
    print("Press Ctrl+C to stop the server")
    print("=" * 60 + "\n")

    stcli.main()
    print("\n✓ Streamlit stopped normally.")

def main():
    print("=" * 60)
    print("FRC Ticket GUI - Launcher")
    print("=" * 60)

    if getattr(sys, 'frozen', False):
        print("Mode: Compiled executable")
        base_path = sys._MEIPASS
        app_path = resource_path('app.py')
    else:
        print("Mode: Development script")
        base_path = os.path.dirname(os.path.abspath(__file__))
        app_path = os.path.join(base_path, 'app.py')

    print(f"Base path: {base_path}")
    print(f"App path: {app_path}")

    if not os.path.exists(app_path):
        print(f"\n✗ ERROR: app.py not found at: {app_path}")
        print(f"Current working directory: {os.getcwd()}")
        print(f"Files in base path: {os.listdir(base_path) if os.path.exists(base_path) else 'N/A'}")
        input("\nPress Enter to exit...")
        return

    print("✓ Found app.py")

    # start browser opener
    print("\nStarting browser thread...")
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()

    try:
        run_streamlit(app_path)
    except KeyboardInterrupt:
        print("\n\n✓ Shutting down...")
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()