from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    return """
    <h1>Police Case Diary System</h1>
    <h2>Deployment Successful</h2>
    <p>Your Flask application is now live on Render.</p>
    """