import requests
import json
import win32com.client as wincom
import speech_recognition as sr

speak = wincom.Dispatch("SAPI.SpVoice")
speak.Speak("HELLO,Welcome to my weather app")
r = sr.Recognizer()
my_device_mic = sr.Microphone(device_index=1)
with my_device_mic as source:
    speak.Speak("What's your city")
    r.adjust_for_ambient_noise(source)
    audio = r.listen(source)
my_string = r.recognize_google(audio)
print(my_string)

city = my_string

URL = f"http://api.weatherapi.com/v1/current.json?key=ede2ced97a41423a98395645242611&q={city}"
response = requests.get(URL)
#print(response.text)

currenttemp = json.loads(response.text)
tempc = currenttemp["current"]["temp_c"]
tempcondition = currenttemp["current"]["condition"]["text"]

weatherspeaker = f"The Current weather of{city} is {tempc} degree Celsius and its{tempcondition}"
speak.Speak(weatherspeaker)
if(tempc < 20):
    speak.Speak("Its Cold outside")
elif(tempcondition == "rainy"):
    speak.Speak("Its Rainy outside so don't forget to carry umbrella.")
elif(tempc > 21) and (tempc < 29):
    speak.Speak("Its a good day.")
elif(tempc > 30):
    speak.Speak("Its hot outside.")
speak.Speak('Thankyou')
