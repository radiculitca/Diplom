from dotenv import load_dotenv
import os
from mistralai.client import Mistral

load_dotenv()

api_key = os.getenv("MISTRAL_API_KEY")
if not api_key or api_key == "вставь_сюда_свой_ключ":
    print("Ключ не задан - открой .env и вставь MISTRAL_API_KEY")
    exit(1)

client = Mistral(api_key=api_key)

print("Отправляю тестовый запрос...")
response = client.chat.complete(
    model="mistral-small-latest",
    messages=[{"role": "user", "content": "Напиши одно короткое предложение на русском языке."}]
)

print("OK Ответ:", response.choices[0].message.content)
