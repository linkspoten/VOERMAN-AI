import os, sys, socket
sys.path.insert(0, os.path.dirname(__file__)); sys.path.insert(0, os.getcwd())
from dotenv import load_dotenv
load_dotenv(override=False)

# Ensure out folder
OUT_DIR = os.environ.get('OUT_DIR','out')
os.makedirs(OUT_DIR, exist_ok=True)

# Import app AFTER sys.path is set
from app import app
import uvicorn

def pick_port():
    for p in (8000, 8001, 8181, 8888):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.3)
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return 8000

if __name__ == "__main__":
    port = pick_port()
    # write selected port so the BAT can open the right URL
    with open(os.path.join(OUT_DIR, "port.txt"), "w") as f:
        f.write(str(port))
    print(f"[INFO] Serving on http://127.0.0.1:{port}")
    # Auto-open browser in a background thread
    import threading, time, webbrowser
    def _open():
        for _ in range(20):
            time.sleep(0.5)
            try:
                webbrowser.open(f"http://127.0.0.1:{port}/dashboard")
                break
            except Exception:
                pass
    threading.Thread(target=_open, daemon=True).start()
    try:
        uvicorn.run(app, host="127.0.0.1", port=port, log_level="info", reload=False)
    except Exception as e:
        import traceback; traceback.print_exc(); input("\n[SERVER] Crash details above. Press Enter to close...")
