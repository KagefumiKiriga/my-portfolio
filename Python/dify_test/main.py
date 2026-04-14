import requests
from config import DIFY_API_KEY, DIFY_API_URL, USER_ID

def chat(message):
    response = requests.post(
        DIFY_API_URL,
        headers={
            "Authorization": f"Bearer {DIFY_API_KEY}",
            "Content-Type": "application/json"
        },
        json={
            "inputs": {},
            "query": message,
            "response_mode": "blocking",
            "user": USER_ID
        }
    )
    result = response.json()
    return result["answer"]

if __name__ == "__main__":
    message = "こんにちは！"
    print("送信:", message)
    answer = chat(message)
    print("みこちゃん:", answer)
