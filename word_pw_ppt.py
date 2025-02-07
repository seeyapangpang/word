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

# ✅ 번역 및 예문 생성 함수
def generate_batch_translations(words):
    try:
        words_string = json.dumps(words)  
        client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant. Always respond in JSON format."},
                {"role": "user", "content": f"Provide IPA pronunciation, a list of Korean translations, and a short English sentence for toddlers along with its Korean translation. Here are the words: {words_string}."}
            ]
        )
        output = response.choices[0].message.content.strip()
        parsed_response = json.loads(output)
        return [
            [
                item.get("word", "").strip(),
                f"[ {item.get('ipa', '발음 없음').replace('/', '').strip()} ]" if item.get('ipa') else "발음 없음",
                item.get("korean", "번역 없음").strip(),
                item.get("example", "No example available").strip(),
                item.get("example_korean", "예문 없음").strip()
            ]
            for item in parsed_response.get("translations", [])
        ]
    except Exception as e:
        return [[word, "발음 없음", "번역 없음", "예문 오류", ""] for word in words]

# ✅ 비밀번호 확인 후 실행
if check_password():
    st.title("단어 번역 및 예문 생성기")
    st.write("엑셀 파일을 업로드하면 단어에 대한 IPA 발음, 번역, 예문을 자동 생성합니다.")

    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

    if "result_df" not in st.session_state:
        st.session_state.result_df = None

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, header=None)
        df = df.iloc[1:, :1]
        df.columns = ["Word"]

        word_count = len(df)
        total_tokens, usd_cost, krw_cost, exchange_rate, estimated_time = estimate_cost(word_count)

        st.write("업로드된 데이터:")
        st.write(df)

        st.subheader("예상 비용 및 시간")
        st.write(f"- 예상 토큰 수: {total_tokens}")
        st.write(f"- 예상 비용 (USD): ${usd_cost:.4f}")
        st.write(f"- 예상 비용 (KRW): {krw_cost:,.0f}원 (환율: {exchange_rate:.2f} KRW/USD)")
        st.write(f"- 예상 시간: {estimated_time:.2f} 초")

        if st.button("Go (API 요청 시작)"):
            st.write("번역과 예문을 생성하는 중입니다...")
            translations = generate_batch_translations(df["Word"].tolist())
            st.session_state.result_df = pd.DataFrame(
                translations,
                columns=["Word", "IPA", "Korean", "English Example", "Korean Example"]
            )

    if st.session_state.result_df is not None:
        st.subheader("번역 및 예문 생성 결과")
        st.write(st.session_state.result_df)

        excel_data = write_to_excel(st.session_state.result_df)
        st.download_button(
            label="결과 다운로드 (엑셀)",
            data=excel_data,
            file_name="translated_vocabulary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        pptx_data = write_to_pptx(st.session_state.result_df)
        st.download_button(
            label="결과 다운로드 (파워포인트)",
            data=pptx_data,
            file_name="translated_vocabulary.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
