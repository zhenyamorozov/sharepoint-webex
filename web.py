"""
This web application serves two purposes:
    - process OAuth requests for Webex Integration
    - respond to Webex bot webhooks
"""

# Debug: Check Python environment
print("=== RUNTIME DEBUG INFO ===")
print("Python executable:", sys.executable)
print("Python version:", sys.version)
print("Python path:", sys.path)
print("Current working directory:", os.getcwd())

import os
from dotenv import load_dotenv
import requests

from flask import Flask


# load env variables
load_dotenv(override=True)

# load env vars for Flask
FLASK_SECRET_KEY = os.getenv('FLASK_SECRET_KEY', "dev")
if FLASK_SECRET_KEY == "dev":
    print("Flask secret key is not set in env. Set to any random string with FLASK_SECRET_KEY environment variable.")

# verify mandatory env vars for Webex integration
if not os.getenv('WEBEX_INTEGRATION_CLIENT_ID'):
    print("Webex Integration Client ID is missing. Provide with WEBEX_INTEGRATION_CLIENT_ID environment variable.")
    raise SystemExit()
if not os.getenv('WEBEX_INTEGRATION_CLIENT_SECRET'):
    print("Webex Integration Client Secret is missing. Provide with WEBEX_INTEGRATION_CLIENT_SECRET environment variable.")
    raise SystemExit()

# verify mandatory env vars for Webex bot
if not os.getenv('WEBEX_BOT_TOKEN'):
    print("Webex bot access token is missing. Provide with WEBEX_BOT_TOKEN environment variable.")
    raise SystemExit()
if not os.getenv('WEBEX_BOT_ROOM_ID'):
    print("Webex bot room ID is missing. Provide with WEBEX_BOT_ROOM_ID environment variable.")
    raise SystemExit()

# verify mandatory env vars for Sharepoint
if not os.getenv('SHAREPOINT_CLIENT_ID'):
    print("Sharepoint Client ID is missing. Provide with SHAREPOINT_CLIENT_ID environment variable.")
    raise SystemExit()
if not os.getenv('SHAREPOINT_CLIENT_SECRET'):
    print("Sharepoint Client Secret is missing. Provide with SHAREPOINT_CLIENT_SECRET environment variable.")
    raise SystemExit()


# Get the public app URL
webAppPublicUrl = ""
# try to get the ngrok URL (dev)
try:
    r = requests.get("http://localhost:4040/api/tunnels",
        timeout=2)
    webAppPublicUrl = r.json()['tunnels'][0]['public_url']
    print("Obtained public URL from ngrok: " + webAppPublicUrl)
except Exception:
    # do nothing
    pass
# try to get the AWS EC2 instance URL (prod)
try:
    # AWS IMDSv2 requires authentication
    r = requests.put("http://169.254.169.254/latest/api/token",
        headers={"X-aws-ec2-metadata-token-ttl-seconds": "21600"},
        timeout=2)
    imdsToken = r.text
    r = requests.get("http://169.254.169.254/latest/meta-data/public-hostname",
        headers={"X-aws-ec2-metadata-token": imdsToken},
        timeout=2)
    if r.text:
        webAppPublicUrl = "https://" + r.text
        print("Obtained public URL from AWS IMDS: " + webAppPublicUrl)
except Exception:
    # do nothing
    pass
# try to get the public URL from env
if os.getenv("WEBAPP_PUBLIC_DOMAIN_NAME"):
    webAppPublicUrl = "https://" + os.getenv("WEBAPP_PUBLIC_DOMAIN_NAME")
    print("Obtained public URL from environment variable: " + webAppPublicUrl)

if not webAppPublicUrl:
    print("Could not get the web app public URL")
    raise SystemExit()

print("Launch Flask app")

app = Flask(__name__)

app.secret_key = FLASK_SECRET_KEY

@app.route("/")
def root():
    """ / handler """
    print("/ requested")
    return "Hey, this is Sharepoint-Webex running on Flask!<br>This application is open source: <a href=\"https://github.com/zhenyamorozov/sharepoint-webex\">github</a>"

import auth
auth.init(webAppPublicUrl)
app.add_url_rule('/auth', view_func=auth.auth)
app.add_url_rule("/callback", view_func=auth.callback, methods=["GET"])

import bot
bot.init(webAppPublicUrl)
app.add_url_rule("/webhook", view_func=bot.webhook, methods=['GET', 'POST'])

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
