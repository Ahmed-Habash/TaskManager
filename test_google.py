import requests
import json
import os

api_key = os.environ.get("GOOGLE__ApiKey")
cx = os.environ.get("GOOGLE__SearchEngineId")

if not api_key or not cx:
    print("Error: Missing environment variables GOOGLE__ApiKey or GOOGLE__SearchEngineId")
    exit(1)

query = "terraria duke fishron"

url = f"https://www.googleapis.com/customsearch/v1?key={api_key}&cx={cx}&q={query}&searchType=image&num=1"
response = requests.get(url)
print(f"Status: {response.status_code}")
if response.status_code == 200:
    data = response.json()
    if "items" in data:
        for item in data["items"]:
            print(f"Title: {item.get('title')}")
            print(f"Link: {item.get('link')}")
    else:
        print("No items found.")
else:
    print(response.text)
