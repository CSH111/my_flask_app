from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)


@app.route("/")
def index():
    print("index come!")
    return render_template('./index.html')


# @app.route("/api/getFile")
# def xx():
#     return "zxzxzxzx"


# @app.route("/my")
# def xx():
#     return "mymymymy"
