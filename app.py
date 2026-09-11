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


def read_uploaded_tables(uploaded_file):
    """Return every table in one uploaded file with source metadata."""
    tables = []
    if uploaded_file.name.lower().endswith(".csv"):
        try:
            frame = pd.read_csv(uploaded_file, encoding="utf-8-sig")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            frame = pd.read_csv(uploaded_file, encoding="cp949")
        tables.append(("CSV", frame))
    else:
        workbook = pd.ExcelFile(uploaded_file)
        for sheet_name in workbook.sheet_names:
            tables.append((sheet_name, pd.read_excel(workbook, sheet_name=sheet_name)))
    return tables


def normalize_columns(frame):
    frame = frame.copy()
    frame.columns = [str(col).strip() for col in frame.columns]
    return frame


excel_tab, doc_tab, mentor_tab = st.tabs(
    ["📊 엑셀 업무 자동화", "📝 보도자료·보고서", "👩‍💼 신입공무원 멘토"]
)

with excel_tab:
    st.subheader("부서별 명단 자동 취합")
    st.write("여러 CSV·엑셀 파일과 모든 시트를 한 표로 합치고, 오류를 점검해 정리된 엑셀을 만듭니다.")
    uploads = st.file_uploader(
        "부서에서 제출한 CSV 또는 XLSX 파일을 모두 선택하세요",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
    )
    st.caption("열 이름이 같은 자료끼리 자동으로 맞춰집니다. 파일마다 열 순서가 달라도 괜찮습니다.")
    if uploads:
        try:
            frames = []
            source_summary = []
            for uploaded in uploads:
                for sheet_name, source_df in read_uploaded_tables(uploaded):
                    source_df = normalize_columns(source_df)
                    source_df.insert(0, "출처시트", sheet_name)
                    source_df.insert(0, "출처파일", uploaded.name)
                    frames.append(source_df)
                    source_summary.append(
                        {"출처파일": uploaded.name, "출처시트": sheet_name, "행 수": len(source_df)}
                    )
            df = pd.concat(frames, ignore_index=True, sort=False)
            source_summary_df = pd.DataFrame(source_summary)

            st.success(
                f"파일 {len(uploads):,}개 · 표 {len(frames):,}개 · 총 {len(df):,}행을 취합했습니다."
            )
            with st.expander("파일별 취합 현황", expanded=True):
                st.dataframe(source_summary_df, use_container_width=True, hide_index=True)

            data_columns = [col for col in df.columns if col not in ["출처파일", "출처시트"]]
            st.markdown("#### 1. 점검 기준 선택")
            c1, c2 = st.columns(2)
            duplicate_keys = c1.multiselect(
                "중복 판정 열",
                data_columns,
                help="예: 사번 또는 성명+생년월일. 선택하지 않으면 전체 열이 같은 행을 찾습니다.",
            )
            required_columns = c2.multiselect(
                "필수 입력 열",
                data_columns,
                help="선택한 열이 비어 있는 행을 누락으로 표시합니다.",
            )

            duplicate_subset = duplicate_keys or data_columns
            duplicate_mask = df.duplicated(subset=duplicate_subset, keep=False)
            if required_columns:
                missing_mask = df[required_columns].isna() | df[required_columns].astype(str).apply(
                    lambda col: col.str.strip().eq("")
                )
                missing_rows_mask = missing_mask.any(axis=1)
                missing_cells = int(missing_mask.sum().sum())
            else:
                missing_rows_mask = pd.Series(False, index=df.index)
                missing_cells = 0

            findings = []
            for row_idx, row in df.iterrows():
                for col in data_columns:
                    kind = detect_privacy(row[col])
                    if kind:
                        findings.append(
                            {"취합 행": row_idx + 2, "출처파일": row["출처파일"], "출처시트": row["출처시트"],
                             "열": str(col), "탐지 유형": kind}
                        )
            findings_df = pd.DataFrame(
                findings, columns=["취합 행", "출처파일", "출처시트", "열", "탐지 유형"]
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("전체 취합 행", len(df))
            m2.metric("중복 의심 행", int(duplicate_mask.sum()))
            m3.metric("필수값 누락", missing_cells)
            m4.metric("개인정보 형태", len(findings_df))

            st.markdown("#### 2. 최종 명단 정리")
            remove_duplicates = st.checkbox("중복 행은 첫 번째 자료만 남기기", value=False)
            include_source = st.checkbox("결과 파일에 출처파일·출처시트 표시", value=True)
            default_order = (["출처파일", "출처시트"] if include_source else []) + data_columns
            output_columns = st.multiselect(
                "결과에 포함할 열과 순서",
                list(df.columns),
                default=default_order,
                help="선택한 순서대로 결과 엑셀의 열이 만들어집니다.",
            )
            cleaned_df = df.drop_duplicates(subset=duplicate_subset, keep="first") if remove_duplicates else df.copy()
            if output_columns:
                cleaned_df = cleaned_df[output_columns]

            group_column = st.selectbox("부서별·항목별 집계 기준(선택)", ["집계하지 않음"] + data_columns)
            if group_column == "집계하지 않음":
                group_summary_df = pd.DataFrame(columns=["집계 기준", "건수"])
            else:
                group_summary_df = (
                    cleaned_df.groupby(group_column, dropna=False).size().reset_index(name="건수")
                    .sort_values("건수", ascending=False)
                )

            st.markdown("#### 3. 결과 미리보기")
            st.dataframe(cleaned_df.head(200), use_container_width=True, hide_index=True)
            if group_column != "집계하지 않음":
                st.dataframe(group_summary_df, use_container_width=True, hide_index=True)

            if len(findings_df):
                st.warning("개인정보로 보이는 값이 있습니다. 외부 서비스에 전송하지 마세요.")

            summary_df = pd.DataFrame(
                {"점검 항목": ["업로드 파일", "취합 표", "전체 행", "중복 의심 행", "필수값 누락", "개인정보 형태"],
                 "결과": [len(uploads), len(frames), len(df), int(duplicate_mask.sum()), missing_cells, len(findings_df)]}
            )
            result = to_excel_bytes(
                {"정리된명단": cleaned_df, "점검요약": summary_df, "파일별현황": source_summary_df,
                 "중복의심": df[duplicate_mask], "필수값누락": df[missing_rows_mask],
                 "조건별집계": group_summary_df, "개인정보탐지": findings_df}
            )
            st.download_button(
                "📥 정리된 엑셀 다운로드", result, "부서별_명단_취합결과.xlsx", type="primary"
            )
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
st.caption("프로토타입 v0.2 · 개인정보 탐지는 보조 기능이며 모든 개인정보를 완벽히 식별한다는 보장은 없습니다.")
