import os
import sys
import threading
import subprocess
import webview
import time

STREAMLIT_APP = "main.py"  # Substitua pelo seu arquivo principal Streamlit
STREAMLIT_URL = "http://localhost:8501"

def start_streamlit():
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"

    # Inicia o servidor Streamlit apenas uma vez
    if not is_streamlit_running():
        subprocess.Popen(
            [sys.executable, "-m", "streamlit", "run", STREAMLIT_APP],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        wait_for_streamlit()

def is_streamlit_running():
    """
    Verifica se o Streamlit já está rodando na porta padrão.
    """
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect(("localhost", 8501))
        return True
    except socket.error:
        return False
    finally:
        s.close()

def wait_for_streamlit():
    """
    Aguarda o servidor Streamlit estar ativo antes de iniciar o PyWebView.
    """
    while not is_streamlit_running():
        time.sleep(0.5)

if __name__ == "__main__":
    threading.Thread(target=start_streamlit, daemon=True).start()
    webview.create_window("Conversão Cellebrite", STREAMLIT_URL)
    webview.start()
