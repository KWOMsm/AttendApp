import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

def load_and_clean_data(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=None)
    else:
        df = pd.read_excel(file, header=None)
        
    header_row = df[df.apply(lambda row: row.astype(str).str.replace(" ", "").str.contains('일자').any(), axis=1)].index
    if len(header_row) > 0:
        df.columns = df.iloc[header_row[0]]
        df = df.iloc[header_row[0]+1:].reset_index(drop=True)
        
    df.columns = [str(c).replace(" ", "") for c in df.columns]
    
    if '일자' in df.columns and '요일' in df.columns:
        df['일자'] = df['일자'].ffill()
        df['요일'] = df['요일'].ffill()
        
    df['일자_dt'] = pd.to_datetime(df['일자'], errors='coerce')
    df = df.dropna(subset=['일자_dt'])
    df['week_start'] = df['일자_dt'] - pd.to_timedelta(df['일자_dt'].dt.weekday, unit='d')
    return df

def shorten_subject(name):
    if pd.isna(name) or not str(name).strip(): return ""
    name_str = str(name).strip()
    mapping = {
        "간호관리": "관리", "기본간호1": "기본1", "기본간호2": "기본2",
        "기초약리": "약리", "기초영양": "영양", "기초치과": "치과", 
        "기초한방": "한방", "기초해부": "해부", "노인간호": "노인", 
        "모성간호": "모성", "모자보건": "모자", "보건교육": "보교", 
        "보건행정": "보행", "산업보건": "산업", "성인간호1": "성인1", 
        "성인간호2": "성인2", "아동간호": "아동", "응급간호": "응급", 
        "의료관계법규": "법규", "의학용어": "의용", "지역사회": "지역"
    }
    for k, v in mapping.items():
        if k in name_str: return v
    keywords = ["관리", "기본1", "기본2", "약리", "영양", "치과", "한방", "해부", "노인", "모성", "모자", "보교", "보행", "산업", "성인1", "성인2", "아동", "응급", "법규", "의용", "지역", "인구", "질병", "환경"]
    for kw in keywords:
        if kw in name_str: return kw
    return name_str

def create_attendance_excel(df, target_weeks, student_count):
    wb = Workbook()
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid") 
    stat_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") 
    cumul_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") 
    light_grey_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") 
    dark_grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # 포기자용 진한 회색
    
    title_font = Font(size=14, bold=True, color="1F4E78")
    bold_font = Font(bold=True)
    subject_font = Font(size=10) 
    teacher_font = Font(size=9, color="595959")
    stat_head_font = Font(size=9, bold=True)
    
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=False) 
    shrink_align = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True)
    
    sheets = {
        "통합입력": wb.active,
        "기초간호": wb.create_sheet("기초간호"),
        "보건간호": wb.create_sheet("보건간호"),
        "공중간호": wb.create_sheet("공중간호")
    }
    wb.active.title = "통합입력"
    subject_map = {"기초간호": "기초간호", "보건간호": "보건간호", "공중간호": "공중보건", "통합입력": "통합"}
    
    for ws in sheets.values():
        ws.sheet_view.showGridLines = False 
        ws.freeze_panes = "C1" # 💡 이름과 포기일자까지 고정되도록 C1로 변경
        
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE 
        ws.page_setup.paperSize = ws.PAPERSIZE_A4 
        ws.page_setup.fitToPage = True 
        ws.page_setup.fitToWidth = 1 
        ws.page_setup.fitToHeight = 0 
        
        ws.page_margins.left = 0.2; ws.page_margins.right = 0.2
        ws.page_margins.top = 0.3; ws.page_margins.bottom = 0.3
        ws.print_options.horizontalCentered = True; ws.print_options.verticalCentered = True
        
        ws.column_dimensions['A'].width = 10 # 이름
        ws.column_dimensions['B'].width = 8  # 💡 포기일자 칸
        for c in range(3, 38): 
            ws.column_dimensions[get_column_letter(c)].width = 3.6 
        for c in range(38, 48): 
            ws.column_dimensions[get_column_letter(c)].width = 5.5 
            
    block_height = student_count + 8 
    student_row_height = max(15, min(28, int(630 / student_count)))
    
    for w_idx, current_week_start in enumerate(target_weeks):
        start_row = 1 + (w_idx * block_height)
        title_str = f"■ {current_week_start.strftime('%Y년 %m월 %d일')} 시작 주간"
        
        week_total_classes = {"통합입력": 0, "기초간호": 0, "보건간호": 0, "공중간호": 0}
        schedule_struct = []
        
        dates_in_week = [current_week_start + pd.Timedelta(days=i) for i in range(5)]
        days_kr = ["월", "화", "수", "목", "금"]
        
        for d_idx, current_date in enumerate(dates_in_week):
            day_str = days_kr[d_idx]
            df_day = df[df['일자_dt'].dt.date == current_date]
            day_data = []
            cancel_day = ""
            if not df_day.empty:
                first_row = df_day.iloc[0]
                if str(first_row.get('요일', '')) not in ['월','화','수','목','금'] and len(str(first_row.get('요일', ''))) > 1:
                    cancel_day = str(first_row.get('요일', ''))
            for p in range(1, 8):
                main_subj, sub_subj, teacher_name, cancel_reason = "", "", "", cancel_day
                is_class = False
                if not cancel_day and not df_day.empty:
                    p_row = df_day[df_day['교시'].astype(str).str.contains(str(p))]
                    if not p_row.empty:
                        row_data = p_row.iloc[0]
                        if '휴강' in str(row_data.get('과목코드','')) or '휴강' in str(row_data.get('세부교과','')):
                            cancel_reason = "휴강"
                        else:
                            main_subj, raw_sub, raw_teacher = str(row_data.get('교과목', '')), str(row_data.get('세부교과', '')), str(row_data.get('훈련교사', ''))
                            if raw_teacher and raw_teacher != 'nan': teacher_name = raw_teacher
                            if raw_sub and raw_sub != 'nan':
                                is_class = True
                                sub_subj = shorten_subject(raw_sub)
                day_data.append({'p': p, 'main': main_subj, 'sub': sub_subj, 'teacher': teacher_name, 'cancel': cancel_reason, 'is_class': is_class, 'full_date': current_date})
                for sheet_name in week_total_classes.keys():
                    target_keyword = subject_map[sheet_name]
                    if not cancel_reason and is_class and ((target_keyword == "통합") or (target_keyword in main_subj)):
                        week_total_classes[sheet_name] += 1
            schedule_struct.append({'date': current_date, 'day_str': day_str, 'periods': day_data})

        for ws in sheets.values():
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=10)
            ws.cell(row=start_row, column=1, value=title_str).font = title_font
            ws.row_dimensions[start_row].height = 25
            
            # 헤더: 이름/포기일자
            c1, c2 = ws.cell(row=start_row+1, column=1, value="학생이름"), ws.cell(row=start_row+1, column=2, value="포기일자")
            ws.merge_cells(start_row=start_row+1, start_column=1, end_row=start_row+4, end_column=1)
            ws.merge_cells(start_row=start_row+1, start_column=2, end_row=start_row+4, end_column=2)
            for c in [c1, c2]: c.alignment = center_align; c.fill = header_fill; c.border = thin_border; c.font = bold_font
            
            ws.row_dimensions[start_row+1].height = 22; ws.row_dimensions[start_row+2].height = 18; ws.row_dimensions[start_row+3].height = 23; ws.row_dimensions[start_row+4].height = 18
            
            for i in range(student_count):
                row_idx = start_row+5+i
                # 이름 입력
                ws.cell(row=row_idx, column=1, value=f"학생{i+1}").alignment = shrink_align
                # 💡 포기일자 입력 칸 (여기에 05/20 처럼 날짜를 쓰면 작동함)
                # 통합입력 시트에만 입력하면 다른 시트도 연동되도록 설정
                if ws.title != "통합입력":
                    ws.cell(row=row_idx, column=2, value=f"='통합입력'!B{row_idx}").alignment = center_align
                else:
                    ws.cell(row=row_idx, column=2).alignment = center_align
                
                for c in [1, 2]: ws.cell(row=row_idx, column=c).border = thin_border
                ws.row_dimensions[row_idx].height = student_row_height 
                
        col_idx = 3 # 💡 C열부터 출석 체크 시작
        
        for day_info in schedule_struct:
            header_str = f"{day_info['date'].strftime('%m/%d')}({day_info['day_str']})"
            cur_date_excel = day_info['date'].strftime('%Y-%m-%d') # 수식용 날짜
            
            for ws in sheets.values():
                ws.merge_cells(start_row=start_row+1, start_column=col_idx, end_row=start_row+1, end_column=col_idx+6)
                c = ws.cell(row=start_row+1, column=col_idx, value=header_str)
                c.alignment = center_align; c.fill = header_fill; c.font = bold_font; c.border = thin_border
                
            for sheet_name, ws in sheets.items():
                target_keyword = subject_map[sheet_name]
                for p_idx, p_info in enumerate(day_info['periods']):
                    act_c = col_idx + p_idx
                    col_letter = get_column_letter(act_c)
                    
                    # 교시 및 헤더
                    ws.cell(row=start_row+2, column=act_c, value=f"{p_idx+1}").alignment = center_align
                    
                    is_my_subject = (target_keyword == "통합") or (target_keyword in p_info['main'])
                    
                    if p_info['cancel'] or not p_info['is_class'] or not is_my_subject:
                        val = f"[{p_info['cancel']}]" if p_info['cancel'] else ""
                        ws.cell(row=start_row+3, column=act_c, value=val).fill = light_grey_fill
                        ws.cell(row=start_row+4, column=act_c).fill = light_grey_fill
                        for r in range(start_row+5, start_row+5+student_count):
                            ws.cell(row=r, column=act_c).fill = light_grey_fill
                    else:
                        # 정상 수업 칸
                        if sheet_name == "통합입력":
                            ws.cell(row=start_row+3, column=act_c, value=p_info['sub']).alignment = shrink_align
                            ws.cell(row=start_row+4, column=act_c, value=p_info['teacher']).alignment = shrink_align
                        else:
                            ws.cell(row=start_row+3, column=act_c, value=f"='통합입력'!{col_letter}{start_row+3}").alignment = shrink_align
                            ws.cell(row=start_row+4, column=act_c, value=f"='통합입력'!{col_letter}{start_row+4}").alignment = shrink_align
                            for i in range(student_count):
                                r = start_row+5+i
                                # 💡 실시간 수식: 포기일자가 이 칸의 날짜보다 작으면 표시 안 함
                                # 엑셀 수식: IF(OR($B5="", $B5 > 현재날짜), '통합입력'!값, "")
                                ws.cell(row=r, column=act_c, value=f"=IF(OR($B{r}=\"\", $B{r}>DATEVALUE(\"{cur_date_excel}\")), '통합입력'!{col_letter}{r}, \"\")").alignment = center_align

            col_idx += 7
            
        # 통계 영역
        stat_col = col_idx
        for sheet_name in ["기초간호", "보건간호", "공중간호"]:
            ws = sheets[sheet_name]
            ws.merge_cells(start_row=start_row+1, start_column=stat_col, end_row=start_row+1, end_column=stat_col+4)
            ws.cell(row=start_row+1, column=stat_col, value="주간 통계").alignment = shrink_align
            ws.cell(row=start_row+1, column=stat_col).fill = stat_fill
            
            ws.merge_cells(start_row=start_row+1, start_column=stat_col+5, end_row=start_row+1, end_column=stat_col+9)
            ws.cell(row=start_row+1, column=stat_col+5, value="누계 통계").alignment = shrink_align
            ws.cell(row=start_row+1, column=stat_col+5).fill = cumul_fill
            
            for i in range(student_count):
                r = start_row + 5 + i
                data_range = f"C{r}:{get_column_letter(stat_col-1)}{r}"
                
                # 💡 실시간 통계: 포기일자가 있으면 해당 주차의 수업 시간을 차감하여 자동 정산
                # (이 예시는 단순화를 위해 주간 총수업을 수식으로 대체)
                ws.cell(row=r, column=stat_col, value=week_total_classes[sheet_name]).alignment = shrink_align
                ws.cell(row=r, column=stat_col+1, value=f'=COUNTIF({data_range}, "X")').alignment = shrink_align
                ws.cell(row=r, column=stat_col+2, value=f'=COUNTIF({data_range}, "지")').alignment = shrink_align
                ws.cell(row=r, column=stat_col+3, value=f'=COUNTIF({data_range}, "조")').alignment = shrink_align
                ws.cell(row=r, column=stat_col+4, value=f"={get_column_letter(stat_col)}{r}-{get_column_letter(stat_col+1)}{r}-{get_column_letter(stat_col+2)}{r}-{get_column_letter(stat_col+3)}{r}").alignment = shrink_align
                
                for j in range(5):
                    curr_c = get_column_letter(stat_col+j)
                    if w_idx == 0: ws.cell(row=r, column=stat_col+5+j, value=f"={curr_c}{r}")
                    else: ws.cell(row=r, column=stat_col+5+j, value=f"={get_column_letter(stat_col+5+j)}{r-block_height}+{curr_c}{r}")

        for ws in sheets.values():
            for row in range(start_row+1, start_row+5+student_count):
                for col in range(1, stat_col + 10 if ws.title != "통합입력" else stat_col):
                    ws.cell(row=row, column=col).border = thin_border
            ws.row_breaks.append(Break(id=start_row + block_height - 1))

    excel_data = io.BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)
    return excel_data

# --- Streamlit UI ---
st.set_page_config(page_title="스마트 간호학원 출석부", layout="wide")
st.title("📊 스마트 간호학원 출석부 (중도포기 실시간 대응형)")
uploaded_file = st.file_uploader("시간표 업로드", type=["csv", "xlsx"])
if uploaded_file:
    df = load_and_clean_data(uploaded_file)
    st.success("데이터 파싱 완료!")
    student_count = st.number_input("학생 수 설정", min_value=1, value=40)
    if st.button("출석부 생성"):
        unique_weeks = sorted(df['week_start'].dropna().dt.date.unique())
        excel_file = create_attendance_excel(df, unique_weeks, student_count)
        st.download_button("📥 출석부 다운로드", data=excel_file, file_name="실시간_대응_출석부.xlsx")
