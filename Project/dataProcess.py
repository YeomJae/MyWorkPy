"""주석 추가""" 

# pandas 클래스 불러오기 
import pandas as pd

# 아래 사용자 함수에 적용되는 변수 선언하기: Dictionary 사용 
SHEET_NAMES = {
    "bank": "은행자료",
    "saer": "회계자료",
    "combined": "시트합치기",
    "pivot_out": "피봇출금",
    "pivot_in": "피봇입금",
    "error_check": "오류확인대상"
}

# 은행자료 데이터프레임을 가공하는 함수 
def preprocess_bank_data(df_bank):
    # 거래일시 수정 후 sorting, 그리고 불필요 컬럼 제거하기 
    df_bank['거래일시'] = df_bank['거래일시'].str.replace('.', '', regex=False).str.slice(0, 8)
    df_bank = df_bank.sort_values(by='거래일시')
    df_bank = df_bank.drop(columns=['잔액(원)','내 통장 표시','적요','처리점','구분'])
    return df_bank

# 회계자료 데이터프레임을 가공하는 함수 
def preprocess_saer_data(df_saer):
    # 일자 수정 후 sorting, 그리고 특정값이 포함된 불필요 행 및 컬럼 제거하기 
    df_saer['일자'] = df_saer['일자'].str.replace('-', '', regex=False).str.slice(0, 8)
    df_saer = df_saer.sort_values(by='일자')
    df_saer = df_saer[~df_saer['일자'].str.contains('전기 이월|월계|누계', na=False)]
    df_saer = df_saer[~df_saer['적요'].str.contains('합계', na=False)]
    df_saer = df_saer.drop(columns=['전표번호','계정명','잔액','회계단위명'])
    return df_saer

# 은행자료, 회계자료 데이터프레임을 하나로 병합하는 함수 
def combine_df_data(df_bank, df_saer):
    # 은행자료의 컬럼명을 회계자료의 컬럼명과 동일하게 변경하기 
    # 병합 시 통일된 컬럼명에 데이터를 나열하기 위함 
    df_bank = df_bank[['거래일시','출금액(원)','입금액(원)']]
    df_bank.columns = ['거래일시','출금','입금']  # 통일된 컬럼명
    # '구분' 컬럼을 첫번째(제일 왼쪽) 위치에 추가하고 그 값을 은행자료 시트 이름으로 넣어주기 
    df_bank.insert(0, '구분', SHEET_NAMES["bank"])  

    # 회계자료의 컬럼명을 은행자료의 컬럼명과 동일하게 변경하기 
    # 병합 시 통일된 컬럼명에 데이터를 나열하기 위함 
    df_saer = df_saer[['일자','대변','차변']]
    df_saer.columns = ['거래일시','출금','입금']  # 통일된 컬럼명
    # '구분' 컬럼을 첫번째(제일 왼쪽) 위치에 추가하고 그 값을 회계자료 시트 이름으로 넣어주기 
    df_saer.insert(0, '구분', SHEET_NAMES["saer"])

    df_comb = pd.concat([df_bank, df_saer], ignore_index=True)
    return df_comb

    return df_pivot_out

# 피벗테이블을 생성해서 입출금액 각각에 대한 차액 계산 후 오류값 표기하기 
def create_pivot_tables(df_combined, bank_label, saer_label):
    df_pivot_out = df_combined.pivot_table(index="거래일시", columns="구분", values="출금", aggfunc="sum", fill_value=0)
    df_pivot_out["출금차액"] = df_pivot_out[bank_label] - df_pivot_out.get(saer_label, 0)
    df_pivot_out["상태"] = df_pivot_out["출금차액"].apply(lambda x: "정상" if x == 0 else "오류")

    df_pivot_in = df_combined.pivot_table(index="거래일시", columns="구분", values="입금", aggfunc="sum", fill_value=0)
    df_pivot_in["입금차액"] = df_pivot_in[bank_label] - df_pivot_in.get(saer_label, 0)
    df_pivot_in["상태"] = df_pivot_in["입금차액"].apply(lambda x: "정상" if x == 0 else "오류")

    return df_pivot_out, df_pivot_in

# 피벗테이블의 각 데이터프레임을 선별적으로 추출하여 병합하는 함수 
def combine_df_pivot_data(df_pivot_out, df_pivot_in):
    # 피봇출금의 컬럼명을 피봇입금의 컬럼명과 동일하게 변경하기 
    # 병합 시 통일된 컬럼명에 데이터를 나열하기 위함 (출금차액 -> 차액) 
    df_pivot_out = df_pivot_out[['거래일시','은행자료','회계자료','출금차액','상태']]
    df_pivot_out.columns = ['거래일시','은행자료','회계자료','차액','상태']
    # 상태 컬럼이 오류값인 행만 추출하기 
    df_pivot_out = df_pivot_out[df_pivot_out['상태'].str.contains('오류', na=False)]
    # '구분' 컬럼을 첫번째(제일 왼쪽) 위치에 추가하고 그 값을 피봇출금 시트 이름으로 넣어주기 
    df_pivot_out.insert(0, '구분', SHEET_NAMES["pivot_out"])

    # 피봇입금의 컬럼명을 피봇출금의 컬럼명과 동일하게 변경하기 
    # 병합 시 통일된 컬럼명에 데이터를 나열하기 위함 (입금차액 -> 차액) 
    df_pivot_in = df_pivot_in[['거래일시','은행자료','회계자료','입금차액','상태']]
    df_pivot_in.columns = ['거래일시','은행자료','회계자료','차액','상태']
    # 상태 컬럼이 오류값인 행만 추출하기 
    df_pivot_in = df_pivot_in[df_pivot_in['상태'].str.contains('오류', na=False)]
    # '구분' 컬럼을 첫번째(제일 왼쪽) 위치에 추가하고 그 값을 피봇입금 시트 이름으로 넣어주기 
    df_pivot_in.insert(0, '구분', SHEET_NAMES["pivot_in"])

    df_pivot_comb = pd.concat([df_pivot_out,df_pivot_in], ignore_index=True)
    return df_pivot_comb

"""
def get_worker(directory: SystemError, search_item):
    workers_file_path = f"{directory}\\workers.xlsx"
    df = pd.read_excel(workers_file_path,header=0)
    df.columns = df.columns.str.strip()

    # 조건에 맞는 행 필터링
    matched_row = df[df['계좌번호'].astype(str) == search_item]
    # 계좌번호가 있는지 확인 후, 해당 행의 이름과 메일 가져오기
    if not matched_row.empty:
        name = matched_row.iloc[0]['이름']
        email = matched_row.iloc[0]['메일']
    else:
        name = None
        email = None
        util.debug_print(f"{search_item}의 정보를 찾을 수 없습니다.", display=log_bool)
    
    result = {"Name" : name, "Email" : email}
    return result
"""
