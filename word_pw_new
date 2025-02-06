import json
import streamlit as st
import pandas as pd
import time
from io import BytesIO
import requests
from openai import OpenAI
from openpyxl.styles import Font

# ✅ 비밀번호 보호 기능
def check_password():
    """비밀번호 입력이 올바른지 확인"""
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

# ✅ OpenAI API 설정
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ✅ 번역 및 예문 생성 함수 (D열 자동 생성 추가)
def generate_batch_translations(words):
    try:
        words_string = json.dumps(words)  
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant. Always respond in the following JSON format: "
                                              '{"translations": [{"word": "<word>", "ipa": "<IPA pronunciation>", "korean": "<korean translations>", "example": "<very short and simple English sentence for 3-4 year old toddlers>", "example_korean": "<korean translation of the example sentence>"}]}'},
                {"role": "user", "content": f"Provide a list of Korean translations, a very short and simple English sentence for 3-4 year old toddlers, and its Korean translation. The sentence should be very easy, simple, and clear. Avoid difficult words. Use very basic grammar. Here are the words: {words_string}."}
            ]
        )
        output = response.choices[0].message.content.strip()

        try:
            parsed_response = json.loads(output)
        except json.JSONDecodeError as e:
            return [[word, "발음 없음", "번역 없음", "예문 없음(예문 번역 없음).", "예문 없음", "예문 번역 없음"] for word in words]

        translations = []
        for item in parsed_response.get("translations", []):
            word = item.get("word", "").strip()
            ipa = item.get("ipa", "발음 없음").strip()
            ipa = ipa.replace("@", "ə")  

            # "a"의 발음기호를 문맥에 맞게 자동 변환
            if word.lower() == "a":
                ipa = "[ ə ]"  

            ipa = f"[ {ipa.replace('/', '').strip()} ]" if ipa != "발음 없음" else "발음 없음"
            korean = item.get("korean", "번역 없음").strip()
            example = item.get("example", "No example available").strip()
            example_korean = item.get("example_korean", "예문 번역 없음").strip()

            # ✅ D열 데이터 생성: "영문예문(한글예문)."
            combined_example = f"{example} ({example_korean})."

            translations.append([word, ipa, korean, combined_example, example, example_korean])

        return translations
    except Exception as e:
        return [[word, "발음 없음", "번역 없음", "예문 없음(예문 번역 없음).", "예문 없음", "예문 번역 없음"] for word in words]

# ✅ 엑셀 파일 저장 함수 (D, E, F열 포함)
def write_to_excel(result_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False)
    return output.getvalue()

# ✅ 비밀번호 확인 후 실행
if check_password():
    st.title("단어 번역 및 예문 생성기 (D열: '영문예문(한글예문).' 형식 추가)")
    st.write("엑셀 파일을 업로드하면 단어에 대한 IPA 발음, 번역, 예문, 예문 번역을 자동 생성합니다.")

    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

    # ✅ 세션 상태에 데이터가 없으면 초기화
    if "result_df" not in st.session_state:
        st.session_state.result_df = None

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, header=None)
        df = df.iloc[1:, :1]  
        df.columns = ["Word"]

        st.write("업로드된 데이터:")
        st.write(df)

        if st.button("Go (API 요청 시작)"):
            start_time = time.time()
            st.write("번역과 예문을 생성하는 중입니다...")
            
            translations = []
            batch_size = 10
            progress_bar = st.progress(0)

            for i in range(0, len(df), batch_size):
                batch_words = df["Word"].iloc[i:i + batch_size].tolist()
                translations.extend(generate_batch_translations(batch_words))
                progress_bar.progress(min((i + batch_size) / len(df), 1.0))

            end_time = time.time()
            execution_time = end_time - start_time
            
            st.write(f"실제 소요 시간: {execution_time:.2f} 초")
            st.session_state.result_df = pd.DataFrame(translations, columns=["Word", "IPA", "Korean", "Combined Example", "Example Sentence", "Example Korean"])

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
