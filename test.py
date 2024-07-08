from polarion import polarion
import requests
from requests.auth import HTTPBasicAuth

# Replace with your actual Polarion server URL and correct endpoint
url = 'https://alm.mahle/polarion/'

# Replace with your actual username and password
username = 'e0161427'
password = 'compiler@aA123'

try:
    session = requests.Session()
    session.auth = HTTPBasicAuth(username, password)
    response = session.get(f'{url}login')
    if response.status_code == 200:
        print('Logged in successfully')
    else:
        print('Failed to log in')


except Exception as e:
    print("An error occurred:", str(e))