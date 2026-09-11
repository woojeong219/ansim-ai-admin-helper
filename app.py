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
    st.subheader("엑셀 업무 자동화")
    work_type = st.radio(
        "처리할 업무를 선택하세요",
        ["파일·시트 합치기", "중복·누락·오류 찾기", "조건별 집계표", "지정 서식 정리", "부서별 명단 취합"],
        horizontal=True,
    )
    descriptions = {
        "파일·시트 합치기": "여러 파일과 엑셀의 모든 시트를 열 이름에 맞춰 하나의 표로 합칩니다.",
        "중복·누락·오류 찾기": "중복 기준, 필수값, 숫자·날짜 형식을 선택해 오류 위치를 찾습니다.",
        "조건별 집계표": "부서·상태·연도 등 원하는 기준별 건수와 합계·평균을 계산합니다.",
        "지정 서식 정리": "필요한 열만 원하는 순서로 배치하고 열 이름 변경·정렬까지 처리합니다.",
        "부서별 명단 취합": "부서별 제출 명단을 합치고 중복·누락을 점검해 최종 명단을 만듭니다.",
    }
    st.info(descriptions[work_type])
    uploads = st.file_uploader(
        "CSV 또는 XLSX 파일을 선택하세요",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        help="파일을 여러 개 선택할 수 있으며 XLSX 파일은 모든 시트를 읽습니다.",
    )

    if uploads:
        try:
            frames, source_summary = [], []
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
            data_columns = [c for c in df.columns if c not in ["출처파일", "출처시트"]]

            st.success(f"파일 {len(uploads):,}개 · 표 {len(frames):,}개 · 총 {len(df):,}행을 읽었습니다.")
            with st.expander("원본 및 파일별 현황", expanded=False):
                st.dataframe(source_summary_df, use_container_width=True, hide_index=True)
                st.dataframe(df.head(100), use_container_width=True, hide_index=True)

            output_df = df.copy()
            sheets = {"파일별현황": source_summary_df}
            output_name = "엑셀_자동화_결과.xlsx"

            if work_type == "파일·시트 합치기":
                include_source = st.checkbox("출처파일·출처시트 열 포함", value=True)
                if not include_source:
                    output_df = output_df[data_columns]
                sheets["통합자료"] = output_df
                st.metric("통합된 전체 행", len(output_df))
                st.dataframe(output_df.head(200), use_container_width=True, hide_index=True)
                output_name = "파일_시트_통합결과.xlsx"

            elif work_type == "중복·누락·오류 찾기":
                left, right = st.columns(2)
                duplicate_keys = left.multiselect("중복 판정 열", data_columns)
                required_columns = right.multiselect("필수 입력 열", data_columns)
                numeric_columns = left.multiselect("숫자 형식이어야 하는 열", data_columns)
                date_columns = right.multiselect("날짜 형식이어야 하는 열", data_columns)

                duplicate_subset = duplicate_keys or data_columns
                duplicate_mask = df.duplicated(subset=duplicate_subset, keep=False)
                missing_mask = pd.DataFrame(False, index=df.index, columns=data_columns)
                if required_columns:
                    missing_mask[required_columns] = (
                        df[required_columns].isna()
                        | df[required_columns].astype(str).apply(lambda col: col.str.strip().eq(""))
                    )
                format_issues = []
                for col in numeric_columns:
                    invalid = pd.to_numeric(df[col], errors="coerce").isna() & df[col].notna()
                    for idx in df.index[invalid]:
                        format_issues.append({"행": idx + 2, "열": col, "오류": "숫자 형식 아님", "입력값": df.at[idx, col]})
                for col in date_columns:
                    invalid = pd.to_datetime(df[col], errors="coerce").isna() & df[col].notna()
                    for idx in df.index[invalid]:
                        format_issues.append({"행": idx + 2, "열": col, "오류": "날짜 형식 아님", "입력값": df.at[idx, col]})
                issue_df = pd.DataFrame(format_issues, columns=["행", "열", "오류", "입력값"])
                missing_rows = df[missing_mask.any(axis=1)]
                m1, m2, m3 = st.columns(3)
                m1.metric("중복 의심 행", int(duplicate_mask.sum()))
                m2.metric("필수값 누락", int(missing_mask.sum().sum()))
                m3.metric("형식 오류", len(issue_df))
                sheets.update({"원본": df, "중복의심": df[duplicate_mask], "필수값누락": missing_rows, "형식오류": issue_df})
                st.dataframe(issue_df if len(issue_df) else pd.DataFrame({"결과": ["형식 오류 없음"]}), use_container_width=True, hide_index=True)
                output_name = "중복_누락_오류_점검결과.xlsx"

            elif work_type == "조건별 집계표":
                group_columns = st.multiselect("집계 기준 열", data_columns, max_selections=3)
                value_column = st.selectbox("계산할 숫자 열", ["건수만 계산"] + data_columns)
                aggregation = st.selectbox("계산 방법", ["합계", "평균", "최대값", "최소값"])
                if group_columns:
                    if value_column == "건수만 계산":
                        summary = df.groupby(group_columns, dropna=False).size().reset_index(name="건수")
                    else:
                        numeric_values = pd.to_numeric(df[value_column], errors="coerce")
                        temp = df[group_columns].copy()
                        temp[value_column] = numeric_values
                        method_map = {"합계": "sum", "평균": "mean", "최대값": "max", "최소값": "min"}
                        summary = temp.groupby(group_columns, dropna=False)[value_column].agg(method_map[aggregation]).reset_index()
                        summary = summary.rename(columns={value_column: f"{value_column}_{aggregation}"})
                    sheets.update({"원본": df, "조건별집계": summary})
                    st.dataframe(summary, use_container_width=True, hide_index=True)
                else:
                    summary = pd.DataFrame({"안내": ["집계 기준 열을 하나 이상 선택하세요."]})
                    sheets.update({"원본": df, "조건별집계": summary})
                    st.warning("집계 기준 열을 선택하면 결과가 표시됩니다.")
                output_name = "조건별_집계표.xlsx"

            elif work_type == "지정 서식 정리":
                selected_columns = st.multiselect("결과에 포함할 열과 순서", data_columns, default=data_columns)
                rename_text = st.text_area(
                    "열 이름 변경(선택)",
                    placeholder="기존열=새열\n예: 담당표시=담당자",
                    help="한 줄에 하나씩 기존 열과 새 열을 = 기호로 연결하세요.",
                )
                sort_column = st.selectbox("정렬 기준", ["정렬하지 않음"] + selected_columns)
                ascending = st.radio("정렬 방향", ["오름차순", "내림차순"], horizontal=True)
                output_df = df[selected_columns].copy() if selected_columns else pd.DataFrame()
                rename_map = {}
                for line in rename_text.splitlines():
                    if "=" in line:
                        old, new = [part.strip() for part in line.split("=", 1)]
                        if old in output_df.columns and new:
                            rename_map[old] = new
                if sort_column != "정렬하지 않음" and sort_column in output_df.columns:
                    output_df = output_df.sort_values(sort_column, ascending=ascending == "오름차순")
                output_df = output_df.rename(columns=rename_map)
                sheets["정리된자료"] = output_df
                st.dataframe(output_df.head(200), use_container_width=True, hide_index=True)
                output_name = "지정서식_정리결과.xlsx"

            else:
                left, right = st.columns(2)
                department_col = left.selectbox("부서 열", data_columns)
                duplicate_keys = right.multiselect("중복 판정 열", data_columns)
                required_columns = left.multiselect("필수 입력 열", data_columns)
                remove_duplicates = right.checkbox("중복은 첫 번째 행만 남기기", value=False)
                duplicate_subset = duplicate_keys or data_columns
                duplicate_mask = df.duplicated(subset=duplicate_subset, keep=False)
                missing_mask = pd.DataFrame(False, index=df.index, columns=data_columns)
                if required_columns:
                    missing_mask[required_columns] = (
                        df[required_columns].isna()
                        | df[required_columns].astype(str).apply(lambda col: col.str.strip().eq(""))
                    )
                output_df = df.drop_duplicates(subset=duplicate_subset, keep="first") if remove_duplicates else df.copy()
                department_summary = output_df.groupby(department_col, dropna=False).size().reset_index(name="인원")
                sheets.update({"정리된명단": output_df, "부서별현황": department_summary,
                               "중복의심": df[duplicate_mask], "필수값누락": df[missing_mask.any(axis=1)]})
                m1, m2, m3 = st.columns(3)
                m1.metric("전체 명단", len(output_df))
                m2.metric("중복 의심 행", int(duplicate_mask.sum()))
                m3.metric("필수값 누락", int(missing_mask.sum().sum()))
                st.dataframe(department_summary, use_container_width=True, hide_index=True)
                output_name = "부서별_명단_취합결과.xlsx"

            st.download_button("📥 결과 엑셀 다운로드", to_excel_bytes(sheets), output_name, type="primary")
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
st.caption("프로토타입 v0.3 · 개인정보 탐지는 보조 기능이며 모든 개인정보를 완벽히 식별한다는 보장은 없습니다.")
