import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border
from copy import copy
import io

# --- 설정 부분 ---
# 사용자는 이 부분에서 파일 이름만 실제 파일에 맞게 수정하면 됩니다.
FILENAME_CONFIG = {
    "template": "시험결과 통합 양식.xlsx",
    "component": "API-X56L2-D 성분시험결과.xlsx",
    "tensile": "API-X56L2-D 인장시험결과.xlsx",
    "impact": "API-X56L2-D 충격시험결과.xlsx",
    "output": "통합_시험_결과_완성본.xlsx"
}

# --- 데이터 처리 함수들 ---

def get_data(filename, sheet_name=0):
    """엑셀 파일을 안전하게 읽어오는 함수"""
    try:
        return pd.read_excel(filename, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"오류: '{filename}' 파일을 찾을 수 없습니다. 스크립트와 같은 폴더에 파일이 있는지 확인하세요.")
        return None
    except Exception as e:
        print(f"오류: '{filename}' 파일을 읽는 중 문제가 발생했습니다: {e}")
        return None
    
def get_impact_data_with_multiheader(filename):
    """[수정] 2줄 헤더를 가진 충격 시험 엑셀 파일을 읽고 컬럼명을 정리하는 함수"""
    try:
        # header=[0, 1] 옵션으로 2줄 헤더를 정확히 읽음
        df = pd.read_excel(filename, header=[0, 1])
        
        # 다중 레벨 컬럼명을 다루기 쉽게 단일 이름으로 변환
        # 예: ('에너지 (J) SIZE 10 보정', '1') -> '에너지 (J) SIZE 10 보정_1'
        new_columns = []
        for col in df.columns:
            # 첫 번째 레벨 이름에서 'Unnamed:' 부분 제거
            level1 = str(col[0]) if 'Unnamed:' not in str(col[0]) else ''
            # 두 번째 레벨 이름
            level2 = str(col[1]) if 'Unnamed:' not in str(col[1]) else ''
            
            # 두 레벨을 조합하여 최종 컬럼명 생성
            if level1 and level2:
                new_columns.append(f"{level1}_{level2}")
            elif level1:
                new_columns.append(level1)
            else:
                new_columns.append(level2)
        
        df.columns = new_columns
        return df

    except FileNotFoundError:
        print(f"오류: '{filename}' 파일을 찾을 수 없습니다.")
        return None
    except Exception as e:
        print(f"오류: '{filename}' 파일을 읽는 중 문제가 발생했습니다: {e}")
        return None

def process_component_data(df):
    """규칙 2: 성분 시험 데이터 처리"""
    if df is None: return pd.DataFrame()
    
    df['배치번호_키'] = df['시편배치'].str[:8]
    
    # 필요한 성분 컬럼 목록 (템플릿 기준)
    comp_cols = ['C', 'Si', 'Mn', 'P', 'S', 'Cu', 'Ni', 'Cr', 'Mo', 'V', 'Nb', 'Ti', 'Alsol', 'Aloxy', 'Al', 'Ca', 'B', 'PCM', 'CEQ']
    # 기본 정보 컬럼 (한 번만 가져옴)
    info_cols = ['생산오더', '제품배치', '제품기호', '외경', '두께', 'Heat No.', '원재료기호', '원재료업체']
    
    processed_data = {}

    for key, group in df.groupby('배치번호_키'):
        # 마지막 2개 행 선택
        last_two = group.tail(2)
        
        # 1. 기본 정보 추출 (첫 번째 행에서만)
        info_data = last_two.iloc[0][info_cols].to_dict()

        # 2. 성분 데이터 추출 및 컬럼명 변경
        row_data = {}
        for i, (idx, row) in enumerate(last_two.iterrows()):
            suffix = f'_{i+1}' # _1, _2
            for col in comp_cols:
                if col in row:
                    row_data[col + suffix] = row[col]
        
        processed_data[key] = {**info_data, **row_data}
        
    result_df = pd.DataFrame.from_dict(processed_data, orient='index')
    return result_df


def process_tensile_data(df):
    """규칙 3: 인장 시험 데이터 처리"""
    if df is None: return pd.DataFrame()

    df['배치번호_키'] = df['시편배치'].str[:8]
    
    # 처리할 방향과 결과 컬럼 정의
    directions = ["Stripe 모재 L방향", "Stripe 모재 T방향", "Stripe 용접"]
    result_cols = ["YS2 STRESS", "TS STRESS", "연신율 EL(%)", "YR(%)"]
    
    all_data = {}

    for key, group in df.groupby('배치번호_키'):
        key_data = {}
        for direction in directions:
            # 방향별 데이터 필터링 및 마지막 1개 선택
            dir_group = group[group['시편 위치/방향'] == direction]
            if not dir_group.empty:
                last_test = dir_group.iloc[-1]
                for col in result_cols:
                    key_data[f"{direction}_{col}"] = last_test[col]
            else:
                # 데이터가 없으면 0으로 채움
                for col in result_cols:
                    key_data[f"{direction}_{col}"] = None
        all_data[key] = key_data
        
    return pd.DataFrame.from_dict(all_data, orient='index')


def process_impact_data(df):
    """[수정] 규칙 4: 충격 시험 데이터 처리 (시험온도 최빈값 적용)"""
    if df is None: return pd.DataFrame()

    # 컬럼명에 접두사가 붙어있을 수 있으므로, 부분 문자열로 컬럼을 찾음
    specimen_col = next((col for col in df.columns if '시편배치' in col), None)
    notch_col = next((col for col in df.columns if 'Notch 위치' in col), None)
    
    # 컬럼 접두사 정의
    temp_col_prefix = '온도(˚C)'
    energy_col_prefix = '에너지(J) SIZE 10보정'

    if not all([specimen_col, notch_col]):
        print("오류: 충격 시험 파일에서 '시편배치', 'Notch 위치' 컬럼을 찾을 수 없습니다.")
        return pd.DataFrame()

    df['배치번호_키'] = df[specimen_col].str[:8]
    locations = ["Base (Transeverse)", "Weld Line", "HAZ"]
    all_data = {}
    
    for key, group in df.groupby('배치번호_키'):
        key_data = {}
        for loc in locations:
            loc_group = group[group[notch_col] == loc]
            
            if not loc_group.empty:
                last_test_row = loc_group.iloc[-1]
                
                # --- [수정된 온도 처리 로직] ---
                # '온도(˚C)' 아래의 1~6번 값을 리스트로 수집
                temp_values = []
                for i in range(1, 7): # 1부터 6까지 확인
                    col_name = f'{temp_col_prefix}_{i}'
                    if col_name in last_test_row and pd.notna(last_test_row[col_name]):
                        temp_values.append(last_test_row[col_name])
                
                # 최빈값 계산
                test_temperature = None
                if temp_values:
                    # mode() 함수는 최빈값을 Series 형태로 반환하므로, 첫 번째 값을 선택
                    mode_series = pd.Series(temp_values).mode()
                    if not mode_series.empty:
                        test_temperature = mode_series.iloc[0]
                # --- [수정된 온도 처리 로직 끝] ---

                # 에너지 값 추출
                val1 = last_test_row.get(f'{energy_col_prefix}_1', None)
                val2 = last_test_row.get(f'{energy_col_prefix}_2', None)
                val3 = last_test_row.get(f'{energy_col_prefix}_3', None)
                
                valid_values = [v for v in [val1, val2, val3] if pd.notna(v) and isinstance(v, (int, float))]
                
                # 계산된 최빈값을 '온도'로 할당
                key_data[f'{loc}_온도'] = test_temperature
                key_data[f'{loc}_1'] = val1
                key_data[f'{loc}_2'] = val2
                key_data[f'{loc}_3'] = val3
                key_data[f'{loc}_Avg'] = sum(valid_values) / len(valid_values) if valid_values else None
            else:
                for col in ['온도', '1', '2', '3', 'Avg']:
                    key_data[f'{loc}_{col}'] = None
        all_data[key] = key_data
    
    return pd.DataFrame.from_dict(all_data, orient='index')


def main():
    """메인 실행 함수"""
    print("--- 데이터 통합 작업을 시작합니다 ---")

    # 각 데이터 파일 로드 및 전처리
    df_comp_raw = get_data(FILENAME_CONFIG['component'])
    df_tens_raw = get_data(FILENAME_CONFIG['tensile'])
    df_impa_raw = get_impact_data_with_multiheader(FILENAME_CONFIG['impact'])

    # 하나라도 파일 로드에 실패하면 중단
    if any(df is None for df in [df_comp_raw, df_tens_raw, df_impa_raw]):
        print("필수 데이터 파일이 없어 작업을 중단합니다.")
        return

    print("1/4: 성분, 인장, 충격 데이터 처리 중...")
    processed_comp = process_component_data(df_comp_raw)
    processed_tens = process_tensile_data(df_tens_raw)
    processed_impa = process_impact_data(df_impa_raw)

    print("2/4: 처리된 데이터 병합 중...")
    # outer join을 통해 모든 키를 포함하도록 병합
    final_df = processed_comp.join(processed_tens, how='outer')
    final_df = final_df.join(processed_impa, how='outer')

    print(final_df.columns)
    
    # NaN 값을 0으로 채우기
    # final_df.fillna(0, inplace=True)
    
    # 인덱스(배치번호_키)를 다시 컬럼으로 변환
    final_df.reset_index(inplace=True)
    final_df.rename(columns={'index': '시편배치'}, inplace=True)

    print("3/4: 엑셀 템플릿 파일에 데이터 쓰는 중...")
    try:
        wb = openpyxl.load_workbook(FILENAME_CONFIG['template'])
        ws = wb.active
    except FileNotFoundError:
        print(f"오류: 템플릿 파일 '{FILENAME_CONFIG['template']}'을 찾을 수 없습니다.")
        return

    # 템플릿의 헤더 순서 가져오기 (데이터 매핑 기준)
    # 템플릿의 데이터가 4행부터 시작하고, 헤더가 3행에 있다고 가정합니다.
    try:
        template_headers = [cell.value for cell in ws[2]]
    except IndexError:
        print("오류: 템플릿 파일의 3번째 행에 헤더가 존재하지 않습니다.")
        return

    # 데이터 쓰기 시작할 행 찾기
    start_row = ws.max_row + 1
    # 서식을 복사할 템플릿 행 (마지막 데이터 행)
    style_template_row = ws.max_row if ws.max_row > 1 else 1

    # 최종 데이터프레임의 컬럼 순서를 템플릿 헤더에 맞게 재정렬
    # 템플릿에 없는 컬럼은 누락, 최종 데이터에 없는 컬럼은 빈 값으로 처리
    column_map = {
        '시편배치':'시편배치', '생산오더':'생산오더', '제품배치':'제품배치', '제품기호':'제품기호', '외경':'외경', 
        '두께':'두께', 'Heat No.':'Heat No.', '원재료기호':'원재료기호', '원재료업체':'원재료업체',
        'C_1':'C', 'Si_1':'Si', 'Mn_1':'Mn', 'P_1':'P', 'S_1':'S', 'Cu_1':'Cu', 'Ni_1':'Ni', 'Cr_1':'Cr', 
        'Mo_1':'Mo', 'V_1':'V', 'Nb_1':'Nb', 'Ti_1':'Ti', 'Alsol_1':'Alsol', 'Aloxy_1':'Aloxy', 'Al_1':'Al', 
        'Ca_1':'Ca', 'B_1':'B', 'PCM_1':'PCM', 'CEQ_1':'CEQ',
        'C_2':'C', 'Si_2':'Si', 'Mn_2':'Mn', 'P_2':'P', 'S_2':'S', 'Cu_2':'Cu', 'Ni_2':'Ni', 'Cr_2':'Cr', 
        'Mo_2':'Mo', 'V_2':'V', 'Nb_2':'Nb', 'Ti_2':'Ti', 'Alsol_2':'Alsol', 'Aloxy_2':'Aloxy', 'Al_2':'Al', 
        'Ca_2':'Ca', 'B_2':'B', 'PCM_2':'PCM', 'CEQ_2':'CEQ',
        'Stripe 모재 L방향_YS2 STRESS':'YS2 STRESS', 'Stripe 모재 L방향_TS STRESS':'TS STRESS', 
        'Stripe 모재 L방향_연신율 EL(%)':'연신율 EL(%)', 'Stripe 모재 L방향_YR(%)':'YR(%)',
        'Stripe 모재 T방향_YS2 STRESS':'YS2 STRESS', 'Stripe 모재 T방향_TS STRESS':'TS STRESS', 
        'Stripe 모재 T방향_연신율 EL(%)':'연신율 EL(%)', 'Stripe 모재 T방향_YR(%)':'YR(%)',
        'Stripe 용접_YS2 STRESS':'YS2 STRESS', 'Stripe 용접_TS STRESS':'TS STRESS', 
        'Stripe 용접_연신율 EL(%)':'연신율 EL(%)', 'Stripe 용접_YR(%)':'YR(%)',
        'Base (Transeverse)_온도':'온도', 'Base (Transeverse)_1':'1', 'Base (Transeverse)_2':'2', 'Base (Transeverse)_3':'3', 'Base (Transeverse)_Avg':'Avg',
        'Weld Line_온도':'온도', 'Weld Line_1':'1', 'Weld Line_2':'2', 'Weld Line_3':'3', 'Weld Line_Avg':'Avg',
        'HAZ_온도':'온도', 'HAZ_1':'1', 'HAZ_2':'2', 'HAZ_3':'3', 'HAZ_Avg':'Avg'
    }

    # 데이터 쓰기
    for index, row_data in final_df.iterrows():
        current_row = start_row + index
        # 템플릿의 컬럼 순서대로 값을 채워넣기
        for col_idx, header in enumerate(template_headers, 1):
            # 헤더에 맞는 데이터프레임 컬럼 찾기
            # 복잡한 헤더 구조를 감안하여, 순차적으로 매핑된 컬럼을 찾아 값을 입력
            df_col_name = None
            # 이 부분은 템플릿의 복잡한 헤더를 정확히 파싱해야 하므로,
            # 여기서는 순서 기반으로 단순화하여 값을 입력합니다.
            # 보다 정확한 구현을 위해선 헤더 매핑 규칙이 더 명확해야 합니다.
            # 지금은 생성된 final_df의 순서대로 값을 넣는다고 가정합니다.
            if col_idx -1 < len(row_data):
                 ws.cell(row=current_row, column=col_idx, value=row_data.iloc[col_idx-1])

        # 서식 복사
        for col_num in range(1, ws.max_column + 1):
            template_cell = ws.cell(row=style_template_row, column=col_num)
            new_cell = ws.cell(row=current_row, column=col_num)
            
            if template_cell.has_style:
                new_cell.font = copy(template_cell.font)
                new_cell.border = copy(template_cell.border)
                new_cell.fill = copy(template_cell.fill)
                new_cell.number_format = copy(template_cell.number_format)
                new_cell.protection = copy(template_cell.protection)
                new_cell.alignment = copy(template_cell.alignment)

    print(f"4/4: '{FILENAME_CONFIG['output']}' 파일 저장 중...")
    try:
        wb.save(FILENAME_CONFIG['output'])
        print(f"--- 작업 완료! 결과가 '{FILENAME_CONFIG['output']}' 파일에 저장되었습니다. ---")
    except PermissionError:
        print(f"오류: '{FILENAME_CONFIG['output']}' 파일이 다른 프로그램에서 열려있어 저장할 수 없습니다. 파일을 닫고 다시 시도해주세요.")
    except Exception as e:
        print(f"파일 저장 중 오류가 발생했습니다: {e}")


# --- Streamlit 페이지 설정 ---
st.set_page_config(page_title="시험 결과 통합 자동화 툴", layout="wide")

st.title("🔬 시험 결과 통합 자동화 툴")
st.write("아래 4개의 엑셀 파일을 업로드한 후 버튼을 누르면, 규칙에 따라 데이터를 통합하고 서식을 유지한 최종 결과 파일을 다운로드할 수 있습니다.")


# --- UI 부분 ---
st.subheader("1. 파일 업로드")
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("📂 **양식 파일** (.xlsx)", type=['xlsx'])
    component_file = st.file_uploader("📂 **성분 시험 결과** (.xlsx)", type=['xlsx'])
with col2:
    tensile_file = st.file_uploader("📂 **인장 시험 결과** (.xlsx)", type=['xlsx'])
    impact_file = st.file_uploader("📂 **충격 시험 결과** (.xlsx)", type=['xlsx'])

st.divider()

# 모든 파일이 업로드 되었을 때만 버튼과 결과 섹션 표시
if all([template_file, component_file, tensile_file, impact_file]):
    st.subheader("2. 결과 생성")
    if st.button("🚀 결과 생성 및 다운로드", type="primary", use_container_width=True):
        with st.spinner('데이터를 처리하고 엑셀 파일을 생성하는 중입니다... 잠시만 기다려주세요.'):
            # 1. 파일 읽기
            df_comp_raw = get_data(component_file)
            df_tens_raw = get_data(tensile_file)
            df_impa_raw = get_impact_data_with_multiheader(impact_file)

            # 2. 데이터 처리
            processed_comp = process_component_data(df_comp_raw)
            processed_tens = process_tensile_data(df_tens_raw)
            processed_impa = process_impact_data(df_impa_raw)
            
            # 3. 데이터 병합
            final_df = processed_comp.join(processed_tens, how='outer')
            final_df = final_df.join(processed_impa, how='outer')
            final_df.reset_index(inplace=True)
            final_df.rename(columns={'index': '시편배치'}, inplace=True)

            # 4. 템플릿에 데이터 쓰기
            try:
                wb = openpyxl.load_workbook(template_file)
                ws = wb.active
            except Exception as e:
                st.error(f"템플릿 파일을 여는 중 오류가 발생했습니다: {e}")
                st.stop()

            template_headers = [cell.value for cell in ws[3]] 
            start_row = ws.max_row + 1
            style_template_row = ws.max_row if ws.max_row >= 3 else 3
            
            for index, row_data in final_df.iterrows():
                current_row = start_row + index
                for col_idx, value in enumerate(row_data.values, 1):
                    if pd.isna(value): value = None
                    ws.cell(row=current_row, column=col_idx, value=value)
                for col_num in range(1, len(template_headers) + 1):
                    template_cell = ws.cell(row=style_template_row, column=col_num)
                    if not template_cell: continue
                    new_cell = ws.cell(row=current_row, column=col_num)
                    if template_cell.has_style:
                        new_cell.font = copy(template_cell.font)
                        new_cell.border = copy(template_cell.border)
                        new_cell.fill = copy(template_cell.fill)
                        new_cell.number_format = copy(template_cell.number_format)
                        new_cell.protection = copy(template_cell.protection)
                        new_cell.alignment = copy(template_cell.alignment)

            # 5. 최종 엑셀 파일을 메모리에 저장
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)

            st.session_state.download_ready = True
            st.session_state.output_buffer = output_buffer

        st.success("✅ 처리가 완료되었습니다! 아래 버튼을 눌러 파일을 다운로드하세요.")
    
    # 다운로드 버튼 표시 (파일 처리가 완료된 경우)
    if 'download_ready' in st.session_state and st.session_state.download_ready:
        st.download_button(
            label="📥 '통합_시험_결과_완성본.xlsx' 다운로드",
            data=st.session_state.output_buffer,
            file_name="통합_시험_결과_완성본.xlsx",
            mime="application/vnd.ms-excel",
            use_container_width=True
        )

else:
    st.info("💡 4개의 파일을 모두 업로드하면 결과 생성 버튼이 나타납니다.")