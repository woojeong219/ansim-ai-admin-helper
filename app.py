import io
import re
from datetime import date

import pandas as pd
import streamlit as st


st.set_page_config(page_title="안심 AI 행정도우미", page_icon="🔒", layout="wide")

st.markdown(
    """
    <style>
    .block-container {max-width: 1100px; padding-top: 2rem;}
    .safe-box {background:#eef8f1; border:1px solid #b8dfc3; padding:16px 18px;
      border-radius:12px; margin-bottom:18px; color:#174c2c;}
    .small-note {color:#5d6673; font-size:.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🔒 안심 AI 행정도우미")
st.caption("AI는 활용하되, 개인정보는 보내지 않는다")
st.markdown(
    '<div class="safe-box"><b>개인정보 보호 모드</b><br>'
    "이 시연 버전은 외부 AI API를 사용하지 않습니다. 업로드 파일은 현재 실행 세션에서만 처리하며 "
    "별도로 저장하는 기능을 두지 않았습니다. 공개 배포 환경에서는 실제 개인정보·내부자료 대신 "
    "가상 또는 비식별 자료만 사용하세요.</div>",
    unsafe_allow_html=True,
)


PHONE_RE = re.compile(r"(?<!\d)01[016789][\s.-]?\d{3,4}[\s.-]?\d{4}(?!\d)")
RRN_RE = re.compile(r"(?<!\d)\d{6}\s*[-]?\s*[1-8]\d{6}(?!\d)")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def detect_privacy(value):
    text = str(value)
    kinds = []
    if PHONE_RE.search(text):
        kinds.append("휴대전화")
    if RRN_RE.search(text):
        kinds.append("주민등록번호 형태")
    if EMAIL_RE.search(text):
        kinds.append("이메일")
    return ", ".join(kinds)


def mask_text(text):
    text = RRN_RE.sub("[주민등록번호]", text)
    text = PHONE_RE.sub("[연락처]", text)
    text = EMAIL_RE.sub("[이메일]", text)
    return text


def to_excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, dataframe in sheets.items():
            dataframe.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()


excel_tab, doc_tab, mentor_tab = st.tabs(
    ["📊 엑셀 업무 자동화", "📝 보도자료·보고서", "👩‍💼 신입공무원 멘토"]
)

with excel_tab:
    st.subheader("엑셀 업무 자동화")
    st.write("파일을 현재 세션에서 분석해 중복, 빈칸, 개인정보 형태를 찾아줍니다.")
    upload = st.file_uploader("CSV 또는 XLSX 파일", type=["csv", "xlsx"])
    if upload:
        try:
            if upload.name.lower().endswith(".csv"):
                try:
                    df = pd.read_csv(upload, encoding="utf-8-sig")
                except UnicodeDecodeError:
                    upload.seek(0)
                    df = pd.read_csv(upload, encoding="cp949")
            else:
                df = pd.read_excel(upload)

            st.success(f"{len(df):,}행 · {len(df.columns):,}열을 불러왔습니다.")
            st.dataframe(df.head(100), use_container_width=True)

            duplicate_mask = df.duplicated(keep=False)
            empty_count = int(df.isna().sum().sum())
            findings = []
            for row_idx, row in df.iterrows():
                for col in df.columns:
                    kind = detect_privacy(row[col])
                    if kind:
                        findings.append({"행": row_idx + 2, "열": str(col), "탐지 유형": kind})
            findings_df = pd.DataFrame(findings, columns=["행", "열", "탐지 유형"])

            c1, c2, c3 = st.columns(3)
            c1.metric("중복 의심 행", int(duplicate_mask.sum()))
            c2.metric("빈칸", empty_count)
            c3.metric("개인정보 형태", len(findings_df))

            if len(findings_df):
                st.warning("개인정보로 보이는 값이 있습니다. 외부 서비스에 전송하지 마세요.")
                st.dataframe(findings_df, use_container_width=True)
            else:
                st.info("정규식 검사에서 휴대전화·주민등록번호·이메일 형태가 발견되지 않았습니다.")

            summary_df = pd.DataFrame(
                {"점검 항목": ["전체 행", "전체 열", "중복 의심 행", "빈칸", "개인정보 형태"],
                 "결과": [len(df), len(df.columns), int(duplicate_mask.sum()), empty_count, len(findings_df)]}
            )
            result = to_excel_bytes(
                {"원본": df, "점검요약": summary_df, "중복의심": df[duplicate_mask], "개인정보탐지": findings_df}
            )
            st.download_button("📥 점검 결과 엑셀 다운로드", result, "엑셀_점검결과.xlsx")
        except Exception as exc:
            st.error(f"파일을 처리하지 못했습니다: {exc}")

with doc_tab:
    st.subheader("보도자료·보고서 초안 작성")
    st.write("입력 내용을 먼저 마스킹하고, 외부 전송 없이 행정문서 기본 틀을 만듭니다.")
    doc_type = st.radio("문서 종류", ["보도자료", "업무보고"], horizontal=True)
    title = st.text_input("사업명 또는 제목")
    department = st.text_input("담당 부서", placeholder="예: ○○과")
    schedule = st.text_input("기간·일정", placeholder="예: 2026. 9. 21.~9. 25.")
    details = st.text_area("핵심 내용", height=160, placeholder="목적, 대상, 장소, 추진내용, 성과 등을 적어주세요.")
    if st.button("초안 만들기", type="primary"):
        safe_title, safe_dept, safe_schedule, safe_details = map(
            mask_text, [title, department, schedule, details]
        )
        found = sum(bool(detect_privacy(v)) for v in [title, department, schedule, details])
        if found:
            st.warning(f"입력란 {found}곳에서 개인정보 형태를 찾아 마스킹했습니다.")
        if doc_type == "보도자료":
            draft = f"""제목: {safe_title or '[사업명]'}

{safe_dept or '[담당부서]'}는 {safe_schedule or '[기간]'} {safe_title or '[사업]'}을(를) 추진한다.

이번 사업은 [추진 배경 및 목적]을 위해 마련됐다. 주요 내용은 다음과 같다.

{safe_details or '[주요 내용을 입력하세요.]'}

{safe_dept or '[담당부서]'} 관계자는 “시민이 체감할 수 있도록 사업을 차질 없이 추진하겠다”고 밝혔다.

※ 배포 전 수치·고유명사·인용문·담당자 연락처를 반드시 확인하세요."""
        else:
            draft = f"""□ {safe_title or '[보고 제목]'}

○ 추진배경
  - [업무 추진 필요성과 현황]

○ 추진개요
  - 기    간: {safe_schedule or '[기간]'}
  - 담당부서: {safe_dept or '[담당부서]'}
  - 주요내용: {safe_details or '[주요 내용]'}

○ 검토사항
  - [예산, 협조부서, 관련 규정 및 예상 쟁점]

○ 향후계획
  - [세부 일정과 후속 조치]

※ 결재 전 근거·수치·일정·공개 가능 여부를 반드시 확인하세요."""
        st.text_area("생성된 초안", draft, height=380)
        st.download_button("📥 텍스트 파일 다운로드", draft, f"{doc_type}_초안_{date.today()}.txt")

with mentor_tab:
    st.subheader("신입공무원 멘토링")
    st.write("공개 가능한 일반 절차를 안내하는 시연용 지식도우미입니다.")
    topic = st.selectbox("무엇이 궁금한가요?", ["공문 작성", "예산 지출", "출장 처리", "물품 구매", "민원 응대"])
    question = st.text_input("질문을 적어주세요", placeholder="예: 협조 공문은 어떤 순서로 쓰나요?")
    guides = {
        "공문 작성": ("목적과 수신자를 정한 뒤 제목, 근거, 요청사항, 기한, 붙임 순서로 작성하세요.",
                  ["한 문장에는 한 가지 요청만 담기", "기한과 제출 방법을 구체적으로 표시", "붙임 파일명과 본문 표기가 같은지 확인"]),
        "예산 지출": ("예산과목·집행 가능 여부 확인 → 품의 → 계약 또는 구매 → 검수 → 지출서류 구비 → 지급 순으로 확인하세요.",
                  ["사전 품의 여부 확인", "계약·회계 담당자에게 기관별 절차 확인", "증빙의 일자·금액·상호 일치 확인"]),
        "출장 처리": ("출장 전 명령·승인을 받고, 출장 후에는 기관 기준에 따라 복명과 여비 정산을 진행하세요.",
                  ["출장 목적·장소·기간 정확히 입력", "교통·숙박 증빙 보관", "내부 여비 규정과 정산 기한 확인"]),
        "물품 구매": ("수요와 예산을 확인한 뒤 품의, 계약·구매, 납품, 검수, 대금 지급과 물품 등록 여부를 점검하세요.",
                  ["분할 구매로 보이지 않는지 확인", "비교견적 등 필요한 증빙 확인", "자산·물품 등록 대상인지 확인"]),
        "민원 응대": ("민원 요지를 먼저 확인하고 사실관계, 처리기한, 담당부서, 회신 방법을 명확히 안내하세요.",
                  ["확인되지 않은 내용을 단정하지 않기", "개인정보를 불필요하게 기록하지 않기", "폭언·위협 상황은 기관 대응지침 따르기"]),
    }
    if st.button("멘토 답변 보기"):
        answer, checks = guides[topic]
        st.markdown(f"**업무 절차**  \n{answer}")
        st.markdown("**체크리스트**")
        for check in checks:
            st.checkbox(check, key=f"{topic}-{check}")
        st.info("기관별 규정과 내부 결재선이 다를 수 있으므로 최종 처리는 소속기관의 최신 지침과 담당자에게 확인하세요.")

st.divider()
st.caption("프로토타입 v0.1 · 개인정보 탐지는 보조 기능이며 모든 개인정보를 완벽히 식별한다는 보장은 없습니다.")
