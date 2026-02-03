# check_ip.py
import requests

print("Ton IP publique est :")
try:
    response = requests.get('https://api.ipify.org?format=json', timeout=5)
    ip = response.json()['ip']
    print(f"📍 {ip}")
    print(f"\nAjoute cette IP à MongoDB Atlas :")
    print(f"Network Access → Add IP Address → {ip}/32")
except:
    print("Impossible de déterminer ton IP")
    print("Va sur https://whatismyipaddress.com/")