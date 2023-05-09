import requests

TARGET_URL = 'https://notify-api.line.me/api/notify'
TOKEN = 'VWZ5buWV5eDIhAPcuOZnFdqpaSsopFHNLa3QgCKlefH'

MESSAGE = "testmessage"

response = requests.post(
    TARGET_URL,
    headers={'Authorization': 'Bearer ' + TOKEN},
    data = {
        'message': MESSAGE
    }
)

print(response.text)