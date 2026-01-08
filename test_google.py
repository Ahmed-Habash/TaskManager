import requests
import json

api_key = "AIzaSyAMXgzU7ULWruu7StcotDShSJyJseVTjxE"
cx = "d18f1d40571af426a"
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
