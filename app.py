import io
import re
from datetime import date

import pandas as pd
import streamlit as st

from document_converter import ACCEPTED_TYPES, OPERATIONS, convert_uploads


st.set_page_config(page_title="군포시 AI 업무 도우미", page_icon="🏛️", layout="wide")

st.markdown(
    """
    <style>
    .block-container {max-width: 1100px; padding-top: 2rem;}
    .safe-box {background:#eef8f1; border:1px solid #b8dfc3; padding:16px 18px;
      border-radius:12px; margin-bottom:18px; color:#174c2c;}
    .small-note {color:#5d6673; font-size:.9rem;}
    .example-grid {display:grid; grid-template-columns:minmax(0, 1fr) minmax(0, 1fr);
      gap:16px; align-items:stretch; margin:12px 0 18px;}
    .example-card {padding:18px 20px; border-radius:12px; margin:0;
      box-shadow:0 2px 8px rgba(31,41,55,.06);}
    .example-bad {background:#fff5f5; border:2px solid #ef9a9a;}
    .example-good {background:#f0faf3; border:2px solid #77c58b;}
    .example-label {font-weight:800; font-size:1.02rem; margin-bottom:8px;}
    .example-bad .example-label {color:#b42318;}
    .example-good .example-label {color:#176b34;}
    .example-text {font-size:1rem; line-height:1.65; color:#273142;}
    @media (max-width: 700px) {
      .example-grid {grid-template-columns:1fr;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🏛️ 군포시 AI 업무 도우미")
st.caption("AI는 활용하되, 개인정보는 보내지 않는다")
st.markdown(
    '<div class="safe-box"><b>개인정보 보호 모드</b><br>'
    "이 시연 버전은 외부 AI API를 사용하지 않습니다. 업로드 파일은 현재 실행 세션에서만 처리하며 "
    "별도로 저장하는 기능을 두지 않았습니다. 공개 배포 환경에서는 실제 개인정보·내부자료 대신 "
    "가상 또는 비식별 자료만 사용하세요.<br><b>실제 개인정보나 비공개 행정문서에는 사용하지 마세요.</b></div>",
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


def render_example_pair(bad_text, good_text):
    st.markdown(
        f"""
        <div class="example-grid">
          <div class="example-card example-bad">
            <div class="example-label">✕ 잘못된 예시</div>
            <div class="example-text">{bad_text}</div>
          </div>
          <div class="example-card example-good">
            <div class="example-label">✓ 올바른 예시</div>
            <div class="example-text">{good_text}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


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


def read_uploaded_raw_tables(uploaded_file):
    """Read sheets without assuming that the first row contains column names."""
    tables = []
    if uploaded_file.name.lower().endswith(".csv"):
        try:
            frame = pd.read_csv(uploaded_file, header=None, encoding="utf-8-sig")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            frame = pd.read_csv(uploaded_file, header=None, encoding="cp949")
        tables.append(("CSV", frame))
    else:
        workbook = pd.ExcelFile(uploaded_file)
        for sheet_name in workbook.sheet_names:
            tables.append((sheet_name, pd.read_excel(workbook, sheet_name=sheet_name, header=None)))
    return tables


def detect_header_row(raw_frame, scan_rows=30):
    """Guess the header row by looking for a dense, mostly text row near the top."""
    normalized = raw_frame.map(lambda value: "" if pd.isna(value) else str(value).strip())
    limit = min(len(normalized), scan_rows)
    best_row, best_score = 0, -1
    for row_index in range(limit):
        values = [value for value in normalized.iloc[row_index].tolist() if value]
        if not values:
            continue
        unique_count = len(set(values))
        text_count = sum(not value.replace(",", "").replace(".", "", 1).isdigit() for value in values)
        next_density = 0
        if row_index + 1 < len(normalized):
            next_density = sum(bool(value) for value in normalized.iloc[row_index + 1].tolist())
        score = unique_count * 100 + text_count * 5 + min(next_density, unique_count)
        if score > best_score:
            best_row, best_score = row_index, score
    return best_row


def header_row_preview(raw_frame, row_index):
    values = [str(value).strip() for value in raw_frame.iloc[row_index].tolist() if not pd.isna(value) and str(value).strip()]
    return " | ".join(values[:8]) or "빈 행"


def extract_table_from_header_row(raw_frame, header_row):
    """Use a selected row as column names and return the non-empty rows below it."""
    positions, column_names, name_counts = [], [], {}
    for position, value in enumerate(raw_frame.iloc[header_row].tolist()):
        if pd.isna(value) or not str(value).strip():
            continue
        base_name = str(value).strip()
        name_counts[base_name] = name_counts.get(base_name, 0) + 1
        display_name = base_name if name_counts[base_name] == 1 else f"{base_name} ({name_counts[base_name]})"
        positions.append(position)
        column_names.append(display_name)
    if not column_names:
        return None, []

    extracted = raw_frame.iloc[header_row + 1 :, positions].copy()
    extracted.columns = column_names
    extracted = extracted.dropna(how="all")
    blank_rows = extracted.astype("string").apply(lambda col: col.str.strip().eq("")).all(axis=1)
    extracted = extracted[~blank_rows]
    repeated_header = extracted.astype("string").apply(lambda row: list(row.str.strip()) == column_names, axis=1)
    extracted = extracted[~repeated_header].reset_index(drop=True)
    return extracted, column_names


def normalize_columns(frame):
    frame = frame.copy()
    frame.columns = [str(col).strip() for col in frame.columns]
    return frame


def coerce_numeric(series):
    """Convert ordinary Excel amounts to numbers without treating dates as amounts."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.Series(float("nan"), index=series.index)
    cleaned = (
        series.astype("string")
        .str.replace(",", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce")


excel_tab, doc_tab, mentor_tab, conversion_tab = st.tabs(
    ["📊 엑셀 업무 자동화", "📝 보도자료·보고서", "👩‍💼 신입공무원 멘토", "📁 문서 변환"]
)

with excel_tab:
    st.subheader("엑셀 업무 자동화")
    work_type = st.radio(
        "처리할 업무를 선택하세요",
        ["파일·시트 합치기", "중복·누락·오류 찾기", "조건별 집계표", "지정 서식 정리", "부서별 명단 취합", "엑셀 VBA 코드 적용 매뉴얼", "AI 엑셀 활용 방법"],
        horizontal=True,
    )
    descriptions = {
        "파일·시트 합치기": "여러 파일과 엑셀의 모든 시트를 열 이름에 맞춰 하나의 표로 합칩니다.",
        "중복·누락·오류 찾기": "중복 기준, 필수값, 숫자·날짜 형식을 선택해 오류 위치를 찾습니다.",
        "조건별 집계표": "부서·상태·연도 등 원하는 기준별 건수와 합계·평균을 계산합니다.",
        "지정 서식 정리": "필요한 열만 원하는 순서로 배치하고 열 이름 변경·정렬까지 처리합니다.",
        "부서별 명단 취합": "부서별 제출 명단을 합치고 중복·누락을 점검해 최종 명단을 만듭니다.",
        "엑셀 VBA 코드 적용 매뉴얼": "AI에게 VBA 코드를 요청하고, 받은 코드를 Excel에 안전하게 적용하는 방법을 안내합니다.",
        "AI 엑셀 활용 방법": "AI가 엑셀 업무를 정확히 이해하도록 데이터 구조·결과 화면·실행 조건을 설명하는 방법을 안내합니다.",
    }
    st.info(descriptions[work_type])

    if work_type == "엑셀 VBA 코드 적용 매뉴얼":
        st.markdown("#### VBA가 무엇인가요?")
        st.write(
            "VBA(Visual Basic for Applications)는 데스크톱 Excel에 들어 있는 자동화용 프로그래밍 언어입니다. "
            "사람이 매번 클릭하고 복사하던 작업을 코드로 기록해 한 번에 실행할 수 있습니다."
        )
        v1, v2, v3 = st.columns(3)
        v1.markdown("**이럴 때 사용해요**  \n파일 반복 처리, 시트 분리, 서식 적용처럼 여러 단계를 되풀이할 때")
        v2.markdown("**AI의 역할**  \n업무 설명을 바탕으로 VBA 코드 초안을 작성하고 코드 내용을 설명할 때")
        v3.markdown("**Excel의 역할**  \n검토한 VBA 코드를 실제 데이터가 있는 PC에서 실행할 때")
        st.caption("VBA 자체가 AI는 아닙니다. AI가 VBA 코드 초안을 만들고, 사용자가 검토한 뒤 Excel이 코드를 실행하는 방식입니다.")
        prompt_tab, apply_tab, safety_tab = st.tabs(["① 프롬프트 만들기", "② VBA 적용하기", "③ 오류·보안 확인"])
        with prompt_tab:
            st.markdown("#### AI에게 VBA 코드 요청하기")
            task_examples = {
                "여러 파일 반복 처리": "선택한 여러 엑셀 파일의 첫 번째 시트에서 자료를 읽어 현재 통합문서의 '통합' 시트에 이어 붙인다.",
                "조건별로 시트 분리": "'원본' 시트의 부서 열을 기준으로 부서별 시트를 새로 만들고 해당 행을 복사한다.",
                "반복 서식 자동 적용": "사용 중인 표의 제목행, 테두리, 날짜, 숫자 표시 형식을 지정된 규칙으로 정리한다.",
                "중복·빈칸 일괄 처리": "사번 열의 중복값과 필수 열의 빈칸을 찾아 색으로 표시하고 '오류목록' 시트에 기록한다.",
                "폴더 안 파일 자동 취합": "사용자가 선택한 폴더의 모든 xlsx 파일과 각 파일의 지정 시트를 하나의 표로 합친다.",
                "직접 입력": "",
            }

            def load_vba_task_example():
                selected = st.session_state.get("vba_task_choice", "직접 입력")
                st.session_state["vba_task_detail"] = task_examples[selected]

            if "vba_task_choice" not in st.session_state:
                st.session_state["vba_task_choice"] = "여러 파일 반복 처리"
            if "vba_task_detail" not in st.session_state:
                st.session_state["vba_task_detail"] = task_examples[st.session_state["vba_task_choice"]]

            advanced_task = st.selectbox(
                "하고 싶은 작업",
                list(task_examples.keys()),
                key="vba_task_choice",
                on_change=load_vba_task_example,
            )
            task_detail = st.text_area("업무 설명", key="vba_task_detail", height=110)
            sheet_info = st.text_input("시트명과 주요 열", placeholder="예: 원본 시트 / A열 접수번호, B열 부서, C열 금액")
            special_rule = st.text_area(
                "세부 조건과 예외",
                placeholder="예: 제목행은 1행, 빈 행 제외, 기존 시트는 삭제하지 않기, 결과는 새 시트에 만들기",
                height=90,
            )
            st.caption("업무 설명이나 아래 입력 내용을 수정하면 완성 프롬프트도 즉시 바뀝니다. 별도의 생성 버튼은 필요하지 않습니다.")
            generated_prompt = f"""당신은 Excel VBA 전문가입니다. 아래 업무를 수행하는 VBA 코드를 작성해 주세요.

[업무 목적]
{task_detail or '[수행할 업무를 구체적으로 입력]'}

[통합문서 구조]
{sheet_info or '[시트명과 각 열의 의미를 입력]'}

[세부 조건 및 예외]
{special_rule or '[제목행, 빈칸 처리, 결과 위치, 제외 조건 등을 입력]'}

[필수 작성 조건]
1. Windows용 데스크톱 Excel의 표준 모듈에서 실행 가능한 전체 코드를 제공할 것
2. ActiveWorkbook과 ActiveSheet를 무조건 사용하지 말고 대상 통합문서와 시트를 명확히 지정할 것
3. 원본 데이터를 삭제하거나 덮어쓰지 말고, 결과는 새 시트에 만들 것
4. 같은 이름의 결과 시트가 있을 때의 처리 방법을 코드에 포함할 것
5. 오류 처리와 화면 업데이트 복구 코드를 포함할 것
6. 각 코드 구간에 한국어 주석을 달고, 사용자가 수정할 시트명·열 번호를 맨 위에 모을 것
7. 외부 인터넷 접속, 프로그램 실행, 파일 삭제 기능은 사용하지 말 것
8. 코드 아래에 실행 전 준비사항과 실행 방법을 초보자도 알 수 있게 설명할 것

실제 개인정보 예시는 사용하지 말고 가상의 열 이름과 값으로 설명해 주세요."""
            st.caption("아래 상자의 복사 아이콘을 눌러 기관에서 사용이 허용된 AI에 붙여 넣으세요.")
            st.code(generated_prompt, language=None)
            st.warning("실제 명단이나 개인정보를 붙여 넣지 말고, 시트명·열 이름·가상 예시만 설명하세요.")

            with st.expander("안전한 VBA 예제: 선택 영역의 빈칸 표시"):
                safe_vba = '''Option Explicit

Sub HighlightBlankCells()
    Dim targetRange As Range
    Dim cell As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "먼저 점검할 셀 범위를 선택하세요.", vbInformation
        Exit Sub
    End If

    Set targetRange = Selection
    For Each cell In targetRange.Cells
        If Len(Trim(CStr(cell.Value))) = 0 Then
            cell.Interior.Color = RGB(255, 235, 156)
        End If
    Next cell

    MsgBox "빈칸 표시가 완료되었습니다.", vbInformation
End Sub'''
                st.code(safe_vba, language="vb")
                st.caption("현재 선택한 범위의 빈 셀만 노란색으로 표시하며 파일 삭제나 외부 전송 기능은 없습니다.")

        with apply_tab:
            st.markdown("#### 받은 VBA 코드를 Excel에 적용하는 순서")
            st.markdown(
                """
1. **원본 파일의 복사본을 먼저 만듭니다.** 중요한 자료에는 바로 적용하지 않습니다.
2. Windows 데스크톱 Excel에서 복사본을 열고 `다른 이름으로 저장`을 선택합니다.
3. 파일 형식을 **Excel 매크로 사용 통합 문서(`.xlsm`)**로 저장합니다.
4. 키보드에서 `Alt + F11`을 눌러 VBA 편집기를 엽니다.
5. 왼쪽 프로젝트 창에서 현재 통합문서를 선택한 뒤 `삽입(Insert) → 모듈(Module)`을 누릅니다.
6. 오른쪽의 빈 코드 창에 AI가 작성한 VBA 코드 전체를 붙여 넣습니다.
7. 코드 상단의 **시트명·열 번호·결과 시트명**을 실제 파일 구조에 맞게 수정합니다.
8. `디버그(Debug) → VBAProject 컴파일(Compile)`로 문법 오류를 먼저 확인합니다.
9. Excel 화면으로 돌아가 `Alt + F8`을 누르고 실행할 매크로를 선택해 실행합니다.
10. 결과가 맞는지 행 수·합계·누락 여부를 원본과 비교한 뒤 저장합니다.
                """
            )
            st.info("웹용 Excel에서는 VBA 매크로를 만들거나 실행할 수 없습니다. Windows 데스크톱 Excel을 기준으로 한 안내입니다.")

        with safety_tab:
            st.markdown("#### 실행 전에 반드시 확인하세요")
            st.markdown(
                """
- 코드에 `Kill`, `Shell`, `CreateObject("WScript.Shell")`, 인터넷 주소, 이메일 전송 기능이 있으면 실행을 중단합니다.
- 원본 삭제, 전체 행 삭제, 덮어쓰기 코드가 있는지 확인합니다.
- 처음에는 가상 데이터 5~10행이 든 복사본에서 시험합니다.
- 실행 전후의 전체 행 수와 주요 금액 합계를 비교합니다.
- 출처가 불분명한 매크로 파일의 **콘텐츠 사용** 버튼을 누르지 않습니다.
- 기관의 매크로 실행·보안 정책이 우선이며, 차단된 설정을 임의로 해제하지 않습니다.
                """
            )
            st.markdown("#### 자주 생기는 오류")
            st.dataframe(
                pd.DataFrame([
                    ["매크로 목록에 없음", "코드가 표준 모듈에 있는지, Sub가 Private인지 확인"],
                    ["첨자 사용이 잘못되었습니다", "코드의 시트명이 실제 탭 이름과 같은지 확인"],
                    ["형식이 일치하지 않습니다", "숫자·날짜 열에 문자나 빈칸이 섞였는지 확인"],
                    ["매크로 실행이 차단됨", "기관 보안담당자의 정책과 허용 절차 확인"],
                    ["저장 후 코드가 사라짐", "파일을 .xlsx가 아닌 .xlsm 형식으로 저장했는지 확인"],
                ], columns=["증상", "확인할 사항"]),
                use_container_width=True,
                hide_index=True,
            )
    elif work_type == "AI 엑셀 활용 방법":
        st.warning("실제 개인정보나 비공개 행정문서에는 사용하지 마세요. 열 이름과 가상 예시만으로 업무를 설명하세요.")
        st.markdown("#### AI에게 엑셀 업무를 잘 요청하는 3가지 방법")
        st.write(
            "AI는 사용자의 엑셀 화면을 저절로 알 수 없습니다. 데이터가 어떻게 생겼는지, 결과를 어떻게 만들지, "
            "어떤 환경에서 사용할지를 함께 알려주면 훨씬 정확한 답을 받을 수 있습니다."
        )

        st.markdown("### 1. 데이터 구조를 설명하기")
        st.caption("열이 몇 개인지, 각 열에는 어떤 값이 들어 있는지 알려주세요.")
        render_example_pair(
            "엑셀 파일에서 거래처별로 정리해줘.",
            "<b>A열은 날짜, B열은 거래처명, C열은 품목, D열은 금액</b>이야. "
            "같은 거래처별로 금액을 합산하고 거래 건수도 함께 표시해줘.",
        )
        st.info("파일 전체 대신 `A열은 날짜`, `B열은 부서명`처럼 열의 의미를 설명해도 됩니다.")

        st.markdown("### 2. 원하는 최종 화면 설명하기")
        st.caption("결과 표의 열 구성, 표시 방법, 정렬 순서를 구체적으로 적어주세요.")
        render_example_pair(
            "보기 좋게 정리해줘.",
            "결과를 <b>표</b>로 보여줘. 열 구성은 <b>거래처명 / 합계금액 / 거래건수</b>로 하고, "
            "금액은 1,000단위 쉼표로 표시해줘. 합계금액이 높은 순서로 정렬해줘.",
        )

        st.markdown("### 3. 실행 환경과 제약사항 설명하기")
        st.caption("어디에서 실행할지, 인터넷 사용 여부, 사용자의 숙련도와 금지할 기능을 알려주세요.")
        render_example_pair(
            "프로그램 만들어줘.",
            "<b>인터넷 없이 실행 가능한 HTML 파일</b>로 만들어줘. 파일을 드래그앤드롭하면 "
            "결과가 바로 나와야 하고, 코딩을 전혀 모르는 직장인이 사용할 거야. <b>원본 파일은 변경하지 마.</b>",
        )

        st.markdown("### 세 가지를 합친 완성 요청 예시")
        complete_excel_prompt = """엑셀 업무를 자동화하고 싶어.

[데이터 구조]
A열은 날짜, B열은 거래처명, C열은 품목, D열은 금액이야. 첫 번째 행은 제목이고 실제 데이터는 두 번째 행부터 시작해.

[원하는 결과]
같은 거래처별로 금액을 합산하고 거래 건수도 계산해줘. 결과 표의 열은 거래처명 / 합계금액 / 거래건수 순서로 만들어줘. 금액은 1,000단위 쉼표로 표시하고 합계금액이 높은 순서로 정렬해줘.

[실행 환경과 조건]
인터넷 없이 실행 가능한 HTML 파일로 만들어줘. 사용자가 엑셀 파일을 드래그앤드롭하면 결과가 화면에 나타나고, 정리된 엑셀을 다운로드할 수 있어야 해. 코딩을 전혀 모르는 직장인이 사용할 거야. 원본 파일은 변경하거나 삭제하지 마. 실제 개인정보를 외부로 전송하는 기능은 넣지 마."""
        st.code(complete_excel_prompt, language=None)

        st.markdown("### 내 업무에 맞게 작성하는 복사용 틀")
        blank_excel_prompt = """엑셀 업무를 자동화하고 싶어.

[데이터 구조]
- 시트명:
- 제목이 있는 행:
- A열:
- B열:
- C열:
- 데이터 예외사항(빈칸, 중복, 합계행 등):

[원하는 결과]
- 결과 표의 열 구성:
- 계산하거나 분류할 내용:
- 정렬 기준:
- 숫자·날짜 표시 방법:
- 다운로드할 파일 형식:

[실행 환경과 조건]
- 사용할 프로그램 또는 환경:
- 인터넷 연결 가능 여부:
- 사용자의 코딩 숙련도:
- 원본 파일 보존 여부:
- 반드시 제외할 기능:

실제 개인정보나 비공개 행정문서는 사용하지 않고, 가상의 데이터 예시로 설명해줘."""
        st.caption("아래 상자의 복사 아이콘을 누른 뒤 빈칸을 내 업무에 맞게 채우세요.")
        st.code(blank_excel_prompt, language=None)

        with st.expander("AI에게 보내기 전 확인사항", expanded=True):
            st.markdown(
                """
- 각 열의 이름과 의미를 적었나요?
- 데이터가 시작되는 행과 시트명을 적었나요?
- 원하는 결과 열과 정렬 기준을 적었나요?
- 실행할 환경과 사용자의 숙련도를 적었나요?
- 원본 보존, 인터넷 사용 여부 등 제약사항을 적었나요?
- 실제 개인정보와 비공개 행정자료를 제거했나요?
                """
            )
    else:
        if work_type == "부서별 명단 취합":
            st.markdown("#### 기준 열 자동 인식")
            st.caption("파일을 올리면 표의 시작 행과 열 이름을 자동으로 찾습니다. 결과를 확인한 뒤 직접 바꿀 수도 있습니다.")
        uploads = st.file_uploader(
            "CSV 또는 XLSX 파일을 선택하세요",
            type=["csv", "xlsx"],
            accept_multiple_files=True,
            help="파일을 여러 개 선택할 수 있으며 XLSX 파일은 모든 시트를 읽습니다.",
        )

        if not uploads:
            st.caption("파일을 올리면 선택한 업무의 설정 화면이 나타납니다.")

    if work_type not in ("엑셀 VBA 코드 적용 매뉴얼", "AI 엑셀 활용 방법") and uploads:
        try:
            frames, source_summary = [], []
            extraction_errors = []
            if work_type == "부서별 명단 취합":
                raw_tables = []
                for uploaded in uploads:
                    for sheet_name, raw_df in read_uploaded_raw_tables(uploaded):
                        raw_tables.append((uploaded.name, sheet_name, raw_df))

                st.markdown("#### 1. 파일별 기준 열 확인")
                st.caption("자동으로 선택된 행이 실제 열 이름 행과 다르면 해당 파일의 행 번호를 바꿔주세요.")
                for table_index, (file_name, sheet_name, raw_df) in enumerate(raw_tables):
                    detected_row = detect_header_row(raw_df)
                    row_options = list(range(min(len(raw_df), 30)))
                    selected_header_row = st.selectbox(
                        f"{file_name} · {sheet_name}의 기준 열 행",
                        row_options,
                        index=row_options.index(detected_row),
                        format_func=lambda row, frame=raw_df: f"{row + 1}행 — {header_row_preview(frame, row)}",
                        key=f"header_row_{table_index}_{file_name}_{sheet_name}",
                    )
                    source_df, recognized_columns = extract_table_from_header_row(raw_df, selected_header_row)
                    if source_df is None:
                        extraction_errors.append(
                            {"출처파일": file_name, "출처시트": sheet_name,
                             "확인 결과": "선택한 행이 비어 있음"}
                        )
                        continue
                    source_df.insert(0, "출처시트", sheet_name)
                    source_df.insert(0, "출처파일", file_name)
                    frames.append(source_df)
                    source_summary.append(
                        {"출처파일": file_name, "출처시트": sheet_name,
                         "기준 열 위치": f"{selected_header_row + 1}행",
                         "인식된 열": ", ".join(recognized_columns), "취합 행 수": len(source_df)}
                    )
            else:
                for uploaded in uploads:
                    for sheet_name, source_df in read_uploaded_tables(uploaded):
                        source_df = normalize_columns(source_df)
                        source_df.insert(0, "출처시트", sheet_name)
                        source_df.insert(0, "출처파일", uploaded.name)
                        frames.append(source_df)
                        source_summary.append(
                            {"출처파일": uploaded.name, "출처시트": sheet_name, "행 수": len(source_df)}
                        )
            if not frames:
                raise ValueError("선택한 기준 열 행 아래에서 취합할 자료를 찾지 못했습니다.")
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
                st.markdown("#### 무엇을 확인할지 선택하세요")
                st.caption("열은 엑셀 표의 세로 항목을 뜻합니다. 예: 성명, 접수번호, 금액, 신청일")
                with st.expander("선택 방법 예시", expanded=True):
                    st.markdown(
                        """
- **같은 자료 찾기:** `접수번호`를 선택하면 접수번호가 같은 행을 찾아줍니다.
- **빈칸 찾기:** 반드시 작성해야 하는 `성명`, `부서` 등을 선택합니다.
- **숫자가 아닌 값 찾기:** 합계에 사용하는 `금액`, `수량` 등을 선택합니다.
- **날짜가 아닌 값 찾기:** `신청일`, `처리일` 등을 선택합니다.
                        """
                    )
                left, right = st.columns(2)
                duplicate_keys = left.multiselect(
                    "어떤 항목이 같으면 중복으로 볼까요?",
                    data_columns,
                    help="예: 접수번호를 선택하면 접수번호가 같은 행을 찾습니다. 성명과 생년월일을 함께 선택할 수도 있습니다. 아무것도 선택하지 않으면 모든 항목이 똑같은 행을 찾습니다.",
                )
                required_columns = right.multiselect(
                    "비어 있으면 안 되는 항목",
                    data_columns,
                    help="반드시 내용이 있어야 하는 항목을 선택하세요. 선택한 항목의 빈칸을 찾아줍니다.",
                )
                numeric_columns = left.multiselect(
                    "숫자만 입력되어야 하는 항목",
                    data_columns,
                    help="예: 금액, 수량, 인원. 글자나 잘못된 기호가 섞인 값을 찾아줍니다.",
                )
                date_columns = right.multiselect(
                    "날짜로 입력되어야 하는 항목",
                    data_columns,
                    help="예: 신청일, 처리일. 날짜로 인식할 수 없는 값을 찾아줍니다.",
                )

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
                m1.metric("같은 자료 의심", int(duplicate_mask.sum()))
                m2.metric("빈칸 발견", int(missing_mask.sum().sum()))
                m3.metric("입력 형식 확인 필요", len(issue_df))
                sheets.update({"원본": df, "같은자료확인": df[duplicate_mask], "빈칸확인": missing_rows, "입력형식확인": issue_df})
                st.dataframe(issue_df if len(issue_df) else pd.DataFrame({"결과": ["숫자·날짜 입력 형식에 문제가 없습니다."]}), use_container_width=True, hide_index=True)
                output_name = "중복_누락_오류_점검결과.xlsx"

            elif work_type == "조건별 집계표":
                st.caption("집계 기준 열은 자료를 묶는 기준(예: 날짜·부서), 숫자 열은 계산할 값(예: 금액)입니다.")
                group_columns = st.multiselect(
                    "집계 기준 열",
                    data_columns,
                    max_selections=3,
                    help="같은 값끼리 한 행으로 묶을 열입니다. 예: 날짜, 부서, 처리상태",
                )
                numeric_candidates = []
                for col in data_columns:
                    if col in group_columns:
                        continue
                    converted = coerce_numeric(df[col])
                    original_values = df[col].notna().sum()
                    if original_values and converted.notna().sum() / original_values >= 0.8:
                        numeric_candidates.append(col)
                value_column = st.selectbox(
                    "계산할 숫자 열",
                    ["건수만 계산"] + numeric_candidates,
                    help="집계 기준으로 선택한 열과 날짜·문자 열은 자동으로 제외됩니다.",
                )
                if value_column == "건수만 계산":
                    aggregation = None
                    st.caption("선택한 기준별 자료 개수를 계산합니다.")
                else:
                    aggregation = st.selectbox("계산 방법", ["합계", "평균", "최대값", "최소값"])
                if group_columns:
                    if value_column == "건수만 계산":
                        summary = df.groupby(group_columns, dropna=False).size().reset_index(name="건수")
                    else:
                        numeric_values = coerce_numeric(df[value_column])
                        temp = df[group_columns].copy()
                        calculation_name = f"__calculation_{value_column}"
                        temp[calculation_name] = numeric_values
                        method_map = {"합계": "sum", "평균": "mean", "최대값": "max", "최소값": "min"}
                        summary = (
                            temp.groupby(group_columns, dropna=False)[calculation_name]
                            .agg(method_map[aggregation])
                            .reset_index()
                            .rename(columns={calculation_name: f"{value_column}_{aggregation}"})
                        )
                        excluded_values = int(df[value_column].notna().sum() - numeric_values.notna().sum())
                        if excluded_values:
                            st.warning(f"'{value_column}' 열에서 숫자로 바꿀 수 없는 값 {excluded_values:,}건을 계산에서 제외했습니다.")
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
                st.markdown("#### 2. 결과 열과 순서 선택")
                selected_columns = st.multiselect(
                    "결과에 포함할 열",
                    data_columns,
                    default=data_columns,
                    help="파일들에서 인식한 모든 열이 표시됩니다. 선택한 순서가 결과 엑셀의 열 순서가 됩니다.",
                )
                if not selected_columns:
                    raise ValueError("결과에 포함할 열을 하나 이상 선택해 주세요.")

                st.markdown("#### 3. 정렬 및 연번 설정")
                left, right = st.columns(2)
                sort_columns = left.multiselect(
                    "정렬 기준",
                    selected_columns,
                    help="두 개 이상 선택하면 선택한 순서대로 정렬합니다. 각 파일의 서식에 맞는 열을 자유롭게 고를 수 있습니다.",
                )
                sort_direction = right.radio("정렬 방향", ["오름차순", "내림차순"], horizontal=True)
                include_source = left.checkbox("결과에 출처파일·출처시트 표시", value=False)
                sequence_options = ["다시 매기지 않음"] + selected_columns
                default_sequence = sequence_options.index("연번") if "연번" in sequence_options else 0
                sequence_column = right.selectbox(
                    "1번부터 다시 매길 열",
                    sequence_options,
                    index=default_sequence,
                    help="연번·번호처럼 순서를 표시할 열을 선택하세요. 해당 열이 없으면 다시 매기지 않음을 선택합니다.",
                )

                output_df = df.copy()
                if sort_columns:
                    output_df = output_df.sort_values(
                        sort_columns,
                        ascending=sort_direction == "오름차순",
                        na_position="last",
                        key=lambda series: series.astype("string").str.strip(),
                    )
                output_df = output_df.reset_index(drop=True)
                if sequence_column != "다시 매기지 않음":
                    output_df[sequence_column] = range(1, len(output_df) + 1)
                result_columns = (["출처파일", "출처시트"] if include_source else []) + selected_columns
                output_df = output_df[result_columns]

                missing_mask = output_df[selected_columns].isna() | output_df[selected_columns].astype("string").apply(
                    lambda col: col.str.strip().eq("")
                )
                extraction_errors_df = pd.DataFrame(
                    extraction_errors,
                    columns=["출처파일", "출처시트", "확인 결과"],
                )
                coverage_rows = []
                for source_frame in frames:
                    available = [col for col in source_frame.columns if col not in ["출처파일", "출처시트"]]
                    coverage_rows.append(
                        {"출처파일": source_frame["출처파일"].iloc[0],
                         "출처시트": source_frame["출처시트"].iloc[0],
                         "없는 선택 열": ", ".join([col for col in selected_columns if col not in available]) or "없음"}
                    )
                coverage_df = pd.DataFrame(coverage_rows)
                sheets = {
                    "통합명단": output_df,
                    "파일별취합현황": source_summary_df,
                    "열일치확인": coverage_df,
                    "기준열확인필요": extraction_errors_df,
                    "빈칸확인": output_df[missing_mask.any(axis=1)],
                }
                m1, m2, m3 = st.columns(3)
                m1.metric("정상 취합 파일·시트", len(source_summary_df))
                m2.metric("취합된 전체 행", len(output_df))
                m3.metric("기준 열 확인 필요", len(extraction_errors_df))
                if extraction_errors:
                    st.warning("일부 파일·시트에서 선택한 기준 열 행이 비어 있어 취합에서 제외했습니다.")
                    st.dataframe(extraction_errors_df, use_container_width=True, hide_index=True)
                if (coverage_df["없는 선택 열"] != "없음").any():
                    st.warning("일부 파일에는 선택한 열이 없습니다. 해당 파일의 값은 결과에서 빈칸으로 표시됩니다.")
                    st.dataframe(coverage_df, use_container_width=True, hide_index=True)
                st.dataframe(output_df.head(300), use_container_width=True, hide_index=True)
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

with conversion_tab:
    st.subheader("문서 변환")
    st.warning(
        "실제 개인정보나 비공개 행정문서에는 사용하지 마세요. "
        "업로드 파일은 변환을 위해 웹서버의 임시 공간을 거칩니다."
    )
    operation = st.selectbox("변환 방법", OPERATIONS)

    if operation == "PDF → HWP (지원 안내)":
        st.markdown("#### PDF를 HWP로 바꾸려면")
        st.write(
            "공개 Streamlit 서버에서 PDF를 편집 가능한 HWP로 안정적으로 변환하는 기능은 제공하기 어렵습니다. "
            "HWP는 한컴오피스 전용 형식이며, 공개 서버에서 사용할 수 있는 변환 엔진으로는 원본 모양과 편집 구조를 보장할 수 없습니다."
        )
        st.markdown(
            """
1. 개인정보가 없는 문서는 한컴오피스의 PDF 열기·변환 기능을 사용합니다.
2. 실제 행정문서는 기관 PC의 승인된 한컴오피스에서 변환합니다.
3. 표 편집이 목적이면 이 앱의 `PDF → Excel`, 본문 편집이 목적이면 `PDF → Word`를 먼저 시험합니다.
4. 스캔 PDF라면 문자 인식(OCR)이 먼저 필요합니다.
            """
        )
    else:
        accepted = ACCEPTED_TYPES[operation]
        conversion_mode = "editable"
        if operation in ("PDF → Word", "PDF → Excel", "HWP → Word"):
            st.markdown("#### 변환 방식 선택")
            mode_label = st.radio(
                "변환 후 어떻게 사용하실 건가요?",
                ["원본 모습 우선", "내용 편집 우선"],
                horizontal=True,
                key=f"conversion_mode_{operation}",
            )
            conversion_mode = "layout" if mode_label == "원본 모습 우선" else "editable"
            original_col, editable_col = st.columns(2)
            original_col.info(
                "**원본 모습 우선**\n\n그림·표·배치를 페이지 이미지로 유지합니다. 원본과 비슷하게 보이지만 글자와 표를 직접 수정하기 어렵습니다."
            )
            editable_col.info(
                "**내용 편집 우선**\n\n글자와 표를 편집 가능한 내용으로 추출합니다. 수정하기 쉽지만 복잡한 서식과 그림 위치가 달라질 수 있습니다."
            )
            if operation == "PDF → Excel" and conversion_mode == "layout":
                st.caption("PDF 각 페이지가 Excel의 개별 시트에 이미지로 들어갑니다. 셀 계산과 표 수정은 할 수 없습니다.")
            elif operation in ("PDF → Word", "HWP → Word") and conversion_mode == "layout":
                st.caption("각 페이지가 Word에 이미지로 들어갑니다. 화면은 유지되지만 글자를 선택하거나 수정하기 어렵습니다.")
            else:
                st.caption("단순한 본문과 표는 편집할 수 있습니다. 스캔 문서는 OCR이 없으면 글자로 추출되지 않습니다.")

        uploaded_documents = st.file_uploader(
            "변환할 파일을 올려주세요",
            type=accepted,
            accept_multiple_files=True,
            key=f"converter_{operation}",
        )
        if operation in ("PDF 여러 개 합치기", "이미지 → PDF"):
            st.caption("업로드 목록에 표시된 순서대로 하나의 PDF에 합칩니다.")
        if operation in ("엑셀 → PDF", "Word → PDF", "HWP → PDF", "HWP → Word"):
            st.caption("문서 변환 엔진을 사용합니다. 글꼴·쪽 설정·지원 형식에 따라 원본과 일부 차이가 생길 수 있습니다.")

        if st.button("변환하기", type="primary", key="convert_button"):
            try:
                with st.spinner("문서를 변환하고 있습니다..."):
                    download_name, converted_bytes, converted_mime = convert_uploads(
                        operation, uploaded_documents, conversion_mode=conversion_mode
                    )
                st.success("변환이 완료되었습니다.")
                st.download_button(
                    "📥 변환 결과 다운로드",
                    converted_bytes,
                    download_name,
                    mime=converted_mime,
                    type="primary",
                )
            except Exception as exc:
                st.error(f"변환하지 못했습니다: {exc}")

st.divider()
st.caption("프로토타입 v1.3 · 개인정보 탐지는 보조 기능이며 모든 개인정보를 완벽히 식별한다는 보장은 없습니다.")
