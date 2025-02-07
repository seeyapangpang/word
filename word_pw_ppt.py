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

    # 엑셀 파일 저장 함수
    def write_to_excel(result_df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False)
        return output.getvalue()

    # 파워포인트 생성 함수
    def create_pptx(result_df):
        prs = Presentation()
        for _, row in result_df.iterrows():
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # 빈 슬라이드
            
            # 제목 (단어)
            title = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
            title.text = row["Word"]

            # 내용 (IPA, 번역, 예문)
            content = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(5))
            content.text = f"IPA: {row['IPA']}\n\n한글 번역: {row['Korean']}\n\n예문: {row['English Example']}\n한글 예문: {row['Korean Example']}"

        output = BytesIO()
        prs.save(output)
        return output.getvalue()

    # Streamlit 앱 실행
    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

    if "result_df" not in st.session_state:
        st.session_state.result_df = None

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, header=None)
        df = df.iloc[1:, :1]  
        df.columns = ["Word"]
        st.session_state.result_df = pd.DataFrame(df, columns=["Word", "IPA", "Korean", "English Example", "Korean Example"])

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

        pptx_data = create_pptx(st.session_state.result_df)
        st.download_button(
            label="결과 다운로드 (파워포인트)",
            data=pptx_data,
            file_name="translated_vocabulary.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
