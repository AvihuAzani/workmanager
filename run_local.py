"""
Run app locally with a public cloudflare tunnel.
Usage: python run_local.py
"""
import subprocess, threading, sys, re, os, time, webbrowser
import io

DIR = os.path.dirname(os.path.abspath(__file__))
CF  = os.path.join(DIR, "cloudflared.exe")

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def copy_to_clipboard(text):
    try:
        subprocess.run(["clip"], input=text.encode("utf-8"), check=True)
        return True
    except:
        return False

def start_flask():
    return subprocess.Popen(
        [sys.executable, "chat_server.py"],
        cwd=DIR, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )

def start_tunnel():
    return subprocess.Popen(
        [CF, "tunnel", "--url", "http://localhost:5000"],
        cwd=DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace"
    )

def watch_tunnel(proc):
    for line in proc.stdout:
        m = re.search(r'https://[a-z0-9-]+\.trycloudflare\.com', line)
        if m:
            url = m.group(0)
            copy_to_clipboard(url)
            webbrowser.open(url)
            print("\n" + "="*50)
            print("  >>> הכתובת שלך <<<")
            print("")
            print("  " + url)
            print("")
            print("  הכתובת הועתקה ללוח + נפתחה בדפדפן!")
            print("  שלח לטלפון: Ctrl+V")
            print("="*50)
            print("\n  (Ctrl+C לסגירה)\n")
            sys.stdout.flush()

print("מפעיל Flask...")
flask_proc = start_flask()
time.sleep(2)

print("מפעיל Cloudflare Tunnel...")
tunnel_proc = start_tunnel()

t = threading.Thread(target=watch_tunnel, args=(tunnel_proc,), daemon=True)
t.start()

print("ממתין לכתובת (~10 שניות)...")
sys.stdout.flush()

try:
    tunnel_proc.wait()
except KeyboardInterrupt:
    print("\nסוגר...")
    tunnel_proc.terminate()
    flask_proc.terminate()
