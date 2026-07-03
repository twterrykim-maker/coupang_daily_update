import pandas as pd
import os
import xlwings as xw
import warnings
import time

warnings.filterwarnings('ignore')

# ==============================================================================
# ⚙ 1. 경로 및 설정
# ==============================================================================
MASTER_FILE_PATH = r"C:\Users\taewoong.kim\OneDrive - L'Oréal\General - -KR- LRP MKT_ECOM TEAM\Coupang_2024 UPDATE\00_공급현황(Master file)\DAILY update\Daily Update SO.STOCK.PPM_26.07.xlsx"

FOLDER_GMV   = r"C:\Users\taewoong.kim\OneDrive - L'Oréal\General - -KR- LRP MKT_ECOM TEAM\Coupang_2024 UPDATE\00_공급현황(Master file)\DAILY update\raw_gmv"
FOLDER_STOCK = r"C:\Users\taewoong.kim\OneDrive - L'Oréal\General - -KR- LRP MKT_ECOM TEAM\Coupang_2024 UPDATE\00_공급현황(Master file)\DAILY update\raw_inventory"
FOLDER_ADS   = r"C:\Users\taewoong.kim\OneDrive - L'Oréal\General - -KR- LRP MKT_ECOM TEAM\Coupang_2024 UPDATE\00_공급현황(Master file)\DAILY update\raw_pa"
FOLDER_SI    = r"C:\Users\taewoong.kim\OneDrive - L'Oréal\General - -KR- LRP MKT_ECOM TEAM\Coupang_2024 UPDATE\00_공급현황(Master file)\DAILY update\raw_si"

# 시트 이름 설정 (엑셀 파일과 토씨 하나 틀리지 않고 정확해야 함)
SHEET_NAME_GMV = 'SO RAW Data'
SHEET_NAME_STOCK = 'Stock RAW Data'
SHEET_NAME_ADS = '광고raw'
SHEET_NAME_SI = 'SI raw'

PIVOT_LIST = [
    ("chart board", "피벗 테이블5"),
    ("AD", "피벗 테이블24"),
    ("SI pivot", "피벗 테이블1")
]

# ==============================================================================
# 🚀 2. 함수 정의
# ==============================================================================
def load_data_from_folder(folder_path, file_type='csv', header_row=0, col_limit=None):
    """폴더 내 파일들을 읽어 DataFrame으로 반환"""
    all_data = []
    if not os.path.exists(folder_path): return pd.DataFrame()

    ext = '.csv' if file_type == 'csv' else '.xlsx'
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(ext) and not f.startswith('~$')]
    
    if not files: return pd.DataFrame()

    print(f"📂 폴더 스캔({file_type}): {len(files)}개 파일 발견")
    
    for file in files:
        file_path = os.path.join(folder_path, file)
        try:
            if file_type == 'csv':
                try: df = pd.read_csv(file_path, encoding='utf-8')
                except: df = pd.read_csv(file_path, encoding='cp949', errors='ignore')
            else: 
                df = pd.read_excel(file_path, header=header_row)
            
            # 필요한 열(col_limit)까지만 자르기
            if col_limit: df_trimmed = df.iloc[:, :col_limit]
            else: df_trimmed = df
            all_data.append(df_trimmed)
        except: pass

    if all_data: return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()

def get_sheet_strict(wb, sheet_name):
    """
    시트를 찾되, 엑셀이 바빠서 못 찾으면 최대 3번까지 재시도하는 안전한 함수
    """
    max_retries = 3
    for attempt in range(max_retries):
        try:
            return wb.sheets[sheet_name]
        except:
            if attempt < max_retries - 1:
                print(f"   ⏳ 엑셀 응답 지연... '{sheet_name}' 시트 다시 찾는 중 ({attempt+1}/{max_retries})")
                time.sleep(2)
            else:
                print(f"\n❌ [최종 오류] 시트를 찾을 수 없습니다: '{sheet_name}'")
                try:
                    print(f"   👉 힌트: 현재 엑셀 파일의 시트 목록: {[s.name for s in wb.sheets]}")
                except: pass
                raise Exception(f"Sheet Not Found: {sheet_name}")

def run_full_automation():
    print("🚀 [통합] 쿠팡 데이터 자동화 시작...")

    if not os.path.exists(MASTER_FILE_PATH):
        print("❌ 마스터 파일을 찾을 수 없습니다."); return

    # 1. 데이터 로드
    print("\n📦 데이터 로딩 중...")
    df_gmv = load_data_from_folder(FOLDER_GMV, 'xlsx', header_row=0, col_limit=29)  # A~AC = 29개 컬럼
    df_stock = load_data_from_folder(FOLDER_STOCK, 'csv', col_limit=21)  # 15 -> 21로 수정 (A~U열)
    df_ads = load_data_from_folder(FOLDER_ADS, 'xlsx', header_row=0, col_limit=44)
    df_si = load_data_from_folder(FOLDER_SI, 'xlsx', header_row=1, col_limit=19)

    if df_gmv.empty and df_stock.empty and df_ads.empty and df_si.empty:
        print("💤 업데이트할 파일이 없습니다. 종료합니다."); return

    # 2. 엑셀 실행
    print("\n⏳ 엑셀을 실행합니다...")
    try:
        if len(xw.apps) > 0: app = xw.apps.active
        else: app = xw.App(visible=True)
            
        app.display_alerts = False
        app.screen_updating = False 

        wb = app.books.open(MASTER_FILE_PATH, update_links=3)
        app.activate()
        print("✅ 엑셀 파일 열기 성공!")

        # ---------------------------------------------------------
        # [A] 매출 (GMV) - 덮어쓰기 (Replace)
        # ---------------------------------------------------------
        if not df_gmv.empty:
            print(f"\n[1/4] 매출 처리 중...")
            ws = get_sheet_strict(wb, SHEET_NAME_GMV)
            header_row = 4
            data_start = header_row + 1
            
            last_row = ws.range('B' + str(ws.cells.last_cell.row)).end('up').row
            if last_row >= data_start:
                # [설명] B열(인덱스 2)부터 시작. df_gmv 컬럼 수(29개, B~AD열)만큼만 동적으로 영역 지정하여 삭제 (수식 보호)
                end_col_idx = 2 + df_gmv.shape[1] - 1  # B(2) + 29 - 1 = 30 → AD열
                end_col_char = xw.utils.col_name(end_col_idx)
                ws.range(f'B{data_start}:{end_col_char}{last_row}').clear_contents()
            
            # 새 데이터를 B열 지정 위치부터 붙여넣기 (B ~ AD열, 29개 컬럼)
            ws.range(f'B{data_start}').options(index=False, header=False).value = df_gmv
            print(f"   👉 {len(df_gmv)}건 덮어쓰기 완료")

        # ---------------------------------------------------------
        # [B] 재고 (Stock) - 덮어쓰기 (Replace)
        # ---------------------------------------------------------
        if not df_stock.empty:
            print(f"\n[2/4] 재고 처리 중...")
            ws = get_sheet_strict(wb, SHEET_NAME_STOCK)
            header_row = 2
            data_start = header_row + 1 
            
            last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
            if last_row >= data_start:
                # [설명] A열(인덱스 1)부터 시작. df_stock 컬럼 수만큼만 동적으로 영역 지정하여 삭제 (우측 수식 보호)
                end_col_idx = 1 + df_stock.shape[1] - 1
                end_col_char = xw.utils.col_name(end_col_idx)
                ws.range(f'A{data_start}:{end_col_char}{last_row}').clear_contents()
            
            ws.range(f'A{data_start}').options(index=False, header=False).value = df_stock
            print(f"   👉 {len(df_stock)}건 교체 완료")

        # ---------------------------------------------------------
        # [C] 광고 (Ads) - 덮어쓰기 (Replace)
        # ---------------------------------------------------------
        if not df_ads.empty:
            print(f"\n[3/4] 광고 처리 중...")
            ws = get_sheet_strict(wb, SHEET_NAME_ADS)
            header_row = 1
            data_start = header_row + 1
            
            last_row = ws.range('C' + str(ws.cells.last_cell.row)).end('up').row
            if last_row >= data_start:
                # [설명] C열(인덱스 3)부터 시작. df_ads 컬럼 수만큼만 동적으로 영역 지정하여 삭제 (A, B열 및 우측 수식 보호)
                end_col_idx = 3 + df_ads.shape[1] - 1
                end_col_char = xw.utils.col_name(end_col_idx)
                ws.range(f'C{data_start}:{end_col_char}{last_row}').clear_contents()
            
            ws.range(f'C{data_start}').options(index=False, header=False).value = df_ads
            print(f"   👉 {len(df_ads)}건 덮어쓰기 완료")

        # ---------------------------------------------------------
        # [D] Sell-in (SI) - 누적 (Append)
        # ---------------------------------------------------------
        if not df_si.empty:
            print(f"\n[4/4] Sell-in 처리 중...")
            
            # [설명] 앞선 대량 작업 후 엑셀에게 1초 휴식 부여 (에러 방지)
            time.sleep(1) 
            
            ws = get_sheet_strict(wb, SHEET_NAME_SI)
            header_row = 1
            
            # [설명] 데이터를 삭제하지 않고 마지막 행을 찾음
            last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
            if last_row < header_row: last_row = header_row
            
            # [설명] 기존 데이터의 맨 아래(last_row + 1)에 이어서 누적(Append)하여 붙여넣기
            ws.range(f'A{last_row+1}').options(index=False, header=False).value = df_si
            print(f"   👉 {len(df_si)}건 누적 추가 완료 (덮어쓰기 X)")

        # ---------------------------------------------------------
        # [E] 피벗 테이블 새로고침
        # ---------------------------------------------------------
        print("\n🔄 지정된 피벗 테이블만 새로고침 중...")
        time.sleep(3) # 피벗 갱신 전 대기

        for sheet_name, pivot_name in PIVOT_LIST:
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    target_sheet = wb.sheets[sheet_name]
                    pivot_obj = target_sheet.api.PivotTables(pivot_name)
                    pivot_obj.RefreshTable()
                    print(f"   ✅ '{pivot_name}' 업데이트 완료 ({sheet_name})")
                    break 
                except Exception as e:
                    if attempt < max_retries - 1:
                        print(f"   ⏳ '{pivot_name}' 찾는 중... ({attempt+1}/{max_retries})")
                        time.sleep(2)
                    else:
                        print(f"   ❌ [최종 실패] '{pivot_name}' 찾기 실패 ({sheet_name})")

        print("\n💾 저장 중...")
        wb.save()
        print("🎉 모든 작업 완료! (엑셀 창은 유지됩니다)")
        
    except Exception as e:
        print(f"\n❌ 작업 중단됨: {e}")
        try:
            wb.close()
            if len(xw.apps) == 1: app.quit()
        except: pass

    finally:
        try:
            app.screen_updating = True
            app.display_alerts = True
        except: pass

if __name__ == "__main__":
    run_full_automation()