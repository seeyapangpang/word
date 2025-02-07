import json
import streamlit as st
import pandas as pd
import time
from io import BytesIO
from openai import OpenAI
import requests
from openpyxl.styles import Font

# ✅ 비밀번호 보호 기능
def check_password():
    """비밀번호 입력이 올바른지 확인"""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        password = st.text_input("비밀번호를 입력하세요:", type="password")

        # 비밀번호를 Streamlit Secrets에서 불러오기
        if password == st.secrets["APP_PASSWORD"]:  
            st.session_state["password_correct"] = True
        else:
            st.error("비밀번호가 틀렸습니다.")
            return False
    return True

# ✅ 비밀번호 확인 후 실행
if check_password():
    st.title("단어 번역 및 예문 생성기")
    st.write("엑셀 파일을 업로드하면 단어에 대한 IPA 발음, 번역, 예문을 자동 생성합니다.")

    # ✅ OpenAI API 키를 Secrets에서 가져오기 (보안 강화)
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

    # 실시간 환율 가져오기
    def get_exchange_rate():
        try:
            response = requests.get("https://api.exchangerate-api.com/v4/latest/USD")
            data = response.json()
            return data["rates"].get("KRW", 1300)  # 기본값 1300
        except Exception:
            return 1300  # 오류 발생 시 기본 환율

    # 토큰 및 비용 계산 함수
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

    # 번역 및 예문 생성 함수
    def generate_batch_translations(words):
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
            translations = []
            for item in parsed_response.get("translations", []):
                word = item.get("word", "").strip()
                ipa = item.get("ipa", "발음 없음").strip()
                ipa = ipa.replace("@", "ə")  # 발음 기호 형식 유지

                # "a"의 발음기호를 문맥에 맞게 자동 변환
                if word.lower() == "a":
                    ipa = "[ ə ]"  # 일반적인 문장에서 약한 발음으로 발음됨 (ex: "a cat")
                
                ipa = f"[ {ipa.replace('/', '').strip()} ]" if ipa != "발음 없음" else "발음 없음"
                korean = item.get("korean", "번역 없음").strip()
                example = item.get("example", "No example available").strip()
                example_korean = item.get("example_korean", "예문 없음").strip()

                combined_example = f"{example} ({example_korean})"
                translations.append([word, ipa, korean, combined_example, example, example_korean])

            return translations
        except Exception as e:
            return [[word, "발음 없음", "번역 없음", "예문 오류", "", ""] for word in words]

    # 엑셀 파일 저장 함수
    def write_to_excel(result_df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False)
        return output.getvalue()

    # Streamlit 앱 실행
    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

    # 세션 상태에 데이터가 없으면 초기화
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
            start_time = time.time()
            st.write("번역과 예문을 생성하는 중입니다...")

            translations = []
            batch_size = 10
            progress_bar = st.progress(0)

            for i in range(0, word_count, batch_size):
                batch_words = df["Word"].iloc[i:i + batch_size].tolist()
                translations.extend(generate_batch_translations(batch_words))
                progress_bar.progress(min((i + batch_size) / word_count, 1.0))

            end_time = time.time()
            execution_time = end_time - start_time

            st.write(f"실제 소요 시간: {execution_time:.2f} 초")
            
            st.session_state.result_df = pd.DataFrame(translations, columns=["Word", "IPA", "Korean", "Combined Example", "English Example", "Korean Example"])

    if st.session_state.result_df is not None:
        st.subheader("번역 및 예문 생성 결과")
        st.write(st.session_state.result_df)

        excel_data = write_to_excel(st.session_state.result_df)
        st.download_button(
            label="결과 다운로드 (엑셀)",
            data=excel_data,
            file_name="translated_vocabulary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download-btn"
        )
