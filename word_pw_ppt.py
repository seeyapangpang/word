import json
import streamlit as st
import pandas as pd
import time
from io import BytesIO
from openai import OpenAI
import requests
from openpyxl.styles import Font
from pptx import Presentation
from pptx.util import Inches

# ✅ 비밀번호 보호 기능
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        password = st.text_input("비밀번호를 입력하세요:", type="password")
        if password == st.secrets["APP_PASSWORD"]:  
            st.session_state["password_correct"] = True
        else:
            st.error("비밀번호가 틀렸습니다.")
            return False
    return True

# ✅ 실시간 환율 가져오기
def get_exchange_rate():
    try:
        response = requests.get("https://api.exchangerate-api.com/v4/latest/USD")
        data = response.json()
        return data["rates"].get("KRW", 1300)
    except Exception:
        return 1300

# ✅ 예상 비용 및 시간 계산 함수
def estimate_cost(word_count, avg_example_length=50):
    token_per_word = 2  
    token_per_example = avg_example_length * 1.2  
    total_tokens = word_count * (token_per_word + token_per_example)
    cost_per_1k_tokens = 0.0015  
    usd_cost = (total_tokens / 1000) * cost_per_1k_tokens
    exchange_rate = get_exchange_rate()
    krw_cost = usd_cost * exchange_rate
    estimated_time = word_count * 0.2  
    return total_tokens, usd_cost, krw_cost, exchange_rate, estimated_time

# ✅ 번역 및 예문 생성 함수 (3번 재시도 로직 추가)
def generate_batch_translations(words, client, retries=3):
    for attempt in range(retries):
        try:
            words_string = json.dumps(words)  
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant. Always respond in the following JSON format: "
                                                         '{"translations": [{"word": "<word>", "ipa": "<IPA pronunciation>", "korean": "<korean translations>", "example": "<short English sentence>", "example_korean": "<Korean translation of the example>"}]}'},
                    {"role": "user", "content": f"Provide IPA pronunciation, a list of Korean translations, and a very short English sentence for toddlers along with its Korean translation. Here are the words: {words_string}."}
                ]
            )
            output = response.choices[0].message.content.strip()
            parsed_response = json.loads(output)
            return [[
                item.get("word", "").strip(),
                f"[ {item.get('ipa', '발음 없음').replace('/', '').strip()} ]" if item.get("ipa", "발음 없음") != "발음 없음" else "발음 없음",
                item.get("korean", "번역 없음").strip(),
                f"{item.get('example', 'No example available').strip()} ({item.get('example_korean', '예문 없음').strip()})",
                item.get("example", "No example available").strip(),
                item.get("example_korean", "예문 없음").strip()
            ] for item in parsed_response.get("translations", [])]
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(2)  # 재시도 전 대기
                continue
            else:
                return [[word, "발음 없음", "번역 없음", "예문 오류", "예문 오류", "예문 오류"] for word in words]

# ✅ 배치 크기 유지 (10)
batch_size = 10
