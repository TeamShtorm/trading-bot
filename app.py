import os
import threading
from flask import Flask
import subprocess

app = Flask(__name__)
PORT = int(os.environ.get("PORT", 5000))

@app.route('/')
@app.route('/health')
def health():
    return "OK", 200

def run_bot():
    # Запускаем ТВОЙ bot.py
    subprocess.run(["python", "bot.py"])

if __name__ == "__main__":
    bot_thread = threading.Thread(target=run_bot)
    bot_thread.start()
    app.run(host="0.0.0.0", port=PORT)