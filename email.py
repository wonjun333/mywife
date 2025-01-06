import streamlit as st
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.mime.application import MIMEApplication 

# 이메일 발신 설정
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587


# 이메일 전송 함수
def send_email(email_address, app_password, to_email, subject, body, attachment_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = email_address
        msg['To'] = to_email
        msg['Subject'] = subject

        # 본문 추가
        msg.attach(MIMEText(body, 'plain'))

        # 첨부파일 추가 부분 (수정된 코드)
        with open(attachment_path, 'rb') as attachment:  # PDF 파일을 바이너리 모드로 읽기
            part = MIMEBase('application', 'pdf')  # PDF MIME 타입 설정
            part.set_payload(attachment.read())  # 파일 내용을 읽어서 첨부

        # 파일을 Base64로 인코딩
        encoders.encode_base64(part)

        # 첨부파일의 Content-Disposition 헤더 설정 (파일 이름과 확장자 포함)
        filename = os.path.basename(attachment_path)
        if filename:
            try:
                filename = filename.encode('ascii')
            except UnicodeEncodeError:
                filename = Header(filename, 'utf-8').encode()    # This is the important line
        part.add_header(
            'Content-Disposition',
            f'attachment; filename="{filename}"',
        )

        # 이메일에 첨부
        msg.attach(part)


        # 이메일 전송
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(email_address, app_password)
            server.send_message(msg)
        return True
    except Exception as e:
        return f"Error: {e}"

# Streamlit UI 시작
st.title("강의평가 자동 이메일 전송 시스템")

# PDF 파일 저장 디렉토리
TEMP_DIR = "temp"

# 디렉토리 생성 (존재하지 않을 경우)
# 이 코드는 앱 실행 시 최초에 실행됩니다.
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# 사용자 입력받기
st.sidebar.header("발신자 정보 입력")
email_address = st.sidebar.text_input("발신 이메일 주소", "")
app_password = st.sidebar.text_input("앱 비밀번호", type="password")

st.header("1. Excel 데이터 업로드 또는 입력")
uploaded_excel = st.file_uploader("Excel 파일 업로드", type=["xlsx"])
uploaded_data = None

if uploaded_excel:
    uploaded_data = pd.read_excel(uploaded_excel)
else:
    sample_data = {
        "이름": ["강사1", "강사2"],
        "email주소": ["example1@gmail.com", "example2@gmail.com"]
    }
    st.write("아래 샘플 데이터를 붙여넣거나 수정 후 사용하세요.")
    uploaded_data = st.data_editor(pd.DataFrame(sample_data))

st.header("2. PDF 파일 업로드")
uploaded_pdfs = st.file_uploader(
    "PDF 파일 업로드 (여러 개 가능)", type=["pdf"], accept_multiple_files=True
)

st.header("3. 이메일 제목과 본문 입력")
email_subject_template = st.text_input("이메일 제목 템플릿", "강의평가 결과 전달드립니다.")
email_body_template = st.text_area(
    "이메일 본문 템플릿",
    "안녕하세요 {이름}님,\n\n강의평가 결과를 첨부드립니다. 확인 부탁드립니다.\n\n감사합니다.",
)

if st.button("이메일 전송"):
    if not email_address or not app_password:
        st.error("발신 이메일과 앱 비밀번호를 입력하세요.")
    elif uploaded_data is None or not uploaded_pdfs:
        st.error("Excel 데이터와 PDF 파일을 모두 업로드하세요.")
    else:
        # PDF 이름 매칭 준비
        unmatched_names = []
        pdf_files = {file.name: file for file in uploaded_pdfs}  # PDF 파일 이름 딕셔너리
        results = []

        # 강사 이름 매칭
        for _, row in uploaded_data.iterrows():
            name = row['이름']
            email = row['email주소']

            # PDF 이름에 강사 이름이 포함되어 있는지 확인
            matched_pdf = None
            for pdf_name in pdf_files.keys():
                if name in pdf_name:
                    matched_pdf = pdf_files[pdf_name]
                    break

            if matched_pdf:
                # PDF 파일 저장
                pdf_path = os.path.join(TEMP_DIR, matched_pdf.name)
                with open(pdf_path, "wb") as f:
                    f.write(matched_pdf.getbuffer())

                # 제목과 본문 생성
                subject = email_subject_template.format(이름=name)
                body = email_body_template.format(이름=name)

                # 이메일 전송
                result = send_email(email_address, app_password, email, subject, body, pdf_path)

                if result is True:
                    results.append(f"성공: {name} ({email})")
                else:
                    results.append(f"실패: {name} ({email}) - {result}")
            else:
                unmatched_names.append(name)

        # 결과 출력
        st.subheader("결과")
        for res in results:
            st.write(res)

        if unmatched_names:
            st.warning("다음 강사의 PDF 파일을 찾을 수 없습니다:")
            for name in unmatched_names:
                st.write(name)
