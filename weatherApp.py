import requests
import json
import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")

while True:

    city = str(input("Enter city name or Enter 'q' to Quit: "))
    if city == "q":
        print("Terminating the Weather App!")
        speaker.speak("Terminating the weather app")
        break
    url = f"http://api.weatherapi.com/v1/current.json?key=1f976719c21340a6918110407241705&q={city}"
    try:
        r = requests.get(url)
        print(r.text)
        dic = json.loads(r.text)
        speaker.speak(f"The temperature in {city} is {dic["current"]["temp_c"]} degree celsius and the weather is {dic["current"]["condition"]["text"]}")
        #speaker.speak(r.text)
    except:
        print("Error! Data not found!")
        speaker.speak("Error! Data not found")