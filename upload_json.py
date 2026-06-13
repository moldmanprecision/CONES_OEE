import requests

url = "https://CONES_oee-7878.onrender.com/upload-json"

files = {
    "file": open("output/oee_data.json", "rb")
}

response = requests.post(url, files=files)

print(response.text)