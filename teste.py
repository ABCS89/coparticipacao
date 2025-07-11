from flask import Flask, render_template
from pathlib import Path

app = Flask(__name__, template_folder="templates")

@app.route("/")
def index():
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
