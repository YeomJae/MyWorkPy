# [pandas, util(사용자커스텀1), dataProcess(사용자커스텀2)] 클래스 불러오기
import pandas as pd
import util
import dataProcess as dp

log_bool = True

# 아래 비즈니스 로직(사용자가 새로 정의한 함수)에서 파일로 저장할 때 반영할 시트이름 지어주기: Dictionary 사용
SHEET_NAMES = {
    "bank": "은행자료",
    "saer": "회계자료",
    "combined": "시트합치기",
    "pivot_out": "피봇출금",
    "pivot_in": "피봇입금",
    "error_check": "오류확인대상"
}

# 비즈니스 로직을 함수로 정의하기
def toExcelErp(directory: SystemError, filename):
    
    # 1. 작업용 파일에서 계좌번호를 추출하여 각 변수에 값으로 지정하기
    filename2 = filename.split('_')[1][-6:]
    account_num = filename.split('_')[1]

    # 2. 작업용 및 작업완료 후 파일의 위치를 각 변수에 값으로 지정하기
    file_paths = {
        "bank": f"{directory}/workF/{filename}.xls",
        "saer": f"{directory}/workF/거래처원장 {filename2}.xls"
    }
    output_file_path = f"{directory}/resultF/{account_num}.xlsx"
    output_file_path = output_file_path.replace("workF","resultF")
    output_file_path = util.save_excel_with_seq(output_file_path)

    # 3. 은행자료, 회계자료 시트를 데이터프레임 형태로 읽고, 헤더가 없으니 None으로 표기 및 불필요한 행 제거하기
    df_bank = pd.read_excel(file_paths["bank"], header=None).iloc[6:]
    df_saer = pd.read_excel(file_paths["saer"], header=None).iloc[7:]
    
    # 4. 1차 가공된 위의 데이터프레임을 하나의 엑셀파일의 각 시트별로 저장하되 시트 이름은 맨 위에 선언한 변수값으로 만들어주고, 인덱스랑 헤더는 없애기
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False, header=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False, header=False)
    
    # 5. 1차 저장한 엑셀 파일에서 각각의 시트를 데이터프레임 형태로 읽어오기
    df_bank = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["bank"])
    df_saer = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["saer"])

    # 6. 읽어온 데이터프레임을 dataProcess라는 클래스에 있는 함수를 이용해서 가공하기
    df_bank = dp.preprocess_bank_data(df_bank)
    df_saer = dp.preprocess_saer_data(df_saer)
    
    # 7. 가공한 데이터프레임을 하나로 병합하여 새로운 데이터프레임으로 clipboard에 저장하기: dataProcess 클래스의 함수 이용
    df_combined = dp.combine_df_data(df_bank, df_saer)

    # 8. 병합한 데이터프레임에서 피봇 테이블 생성하고 각 조건에 맞게 지정한 시트이름으로 clipboard에 저장하기
    df_pivot_out, df_pivot_in = dp.create_pivot_tables(df_combined, SHEET_NAMES["bank"], SHEET_NAMES["saer"])

    # 9. clipboard에 저장한 데이터프레임을 각각의 지정 시트로 분류하고 하나의 엑셀파일로 저장하기
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False)
        df_combined.to_excel(writer, sheet_name=SHEET_NAMES["combined"], index=False)
        df_pivot_out.to_excel(writer, sheet_name=SHEET_NAMES["pivot_out"])
        df_pivot_in.to_excel(writer, sheet_name=SHEET_NAMES["pivot_in"])

    # 10. 피벗테이블로 돌린 입출금 시트를 각각 읽어와서 데이터프레임 형태로 clipboard에 저장하기
    df_pivot_out = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["pivot_out"])
    df_pivot_in = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["pivot_in"])

    # 11. 읽어온 데이터프레임을 하나로 병합하여 새로운 데이터프레임으로 clipboard에 저장하기: dataProcess 클래스의 함수 이용
    df_pivot_comb = dp.combine_df_pivot_data(df_pivot_out, df_pivot_in)

    # 12. clipboard에 저장한 데이터프레임을 지정 시트로 분류하고 하나의 엑셀파일로 최종 저장하기
    with pd.ExcelWriter(output_file_path) as writer:
        df_pivot_comb.to_excel(writer, sheet_name=SHEET_NAMES["error_check"], index=False)
