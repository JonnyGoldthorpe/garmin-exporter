import os
import json
import requests
from base64 import b64encode
from nacl import encoding, public

def encrypt(public_key_str, secret_value):
    key = public.PublicKey(public_key_str.encode(), encoding.Base64Encoder())
    box = public.SealedBox(key)
    encrypted = box.encrypt(secret_value.encode())
    return b64encode(encrypted).decode()

def update_secret(token, repo, secret_name, secret_value):
    url = f"https://api.github.com/repos/{repo}/actions/secrets/public-key"
    headers = {"Authorization": f"token {token}"}
    response = requests.get(url, headers=headers)
    pub_key = response.json()
    
    encrypted = encrypt(pub_key["key"], secret_value)
    
    url = f"https://api.github.com/repos/{repo}/actions/secrets/{secret_name}"
    requests.put(url, headers=headers, json={
        "encrypted_value": encrypted,
        "key_id": pub_key["key_id"]
    })
    print(f"Updated {secret_name}")

token = os.environ["GH_PAT"]
repo = "JonnyGoldthorpe/garmin-exporter"

with open(os.path.expanduser("~/.garth/oauth1_token.json")) as f:
    oauth1 = f.read()

with open(os.path.expanduser("~/.garth/oauth2_token.json")) as f:
    oauth2 = f.read()

update_secret(token, repo, "GARMIN_OAUTH1_TOKEN", oauth1)
update_secret(token, repo, "GARMIN_OAUTH2_TOKEN", oauth2)
