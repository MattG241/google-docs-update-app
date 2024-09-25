from flask import Flask
import os

app = Flask(__name__)

@app.route('/')
def home():
    return "Refund and Email App"

if __name__ == '__main__':
    # Use the PORT environment variable or default to 5000
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
