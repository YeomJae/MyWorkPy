import pandas as pd
import util
import dataProcess as dp

log_bool= True

# 시트 이름 설정
SHEET_NAMES = {
    "bank": "은행자료",
    "saer": "회계자료",
    "combined": "시트합치기",
    "pivot_out": "피봇출금",
    "pivot_in": "피봇입금",
    "error_check": "오류확인대상"
}

def toExcelErp(directory: SystemError, filename):
    
    # 계좌 번호 추출 및 파일 경로 설정
    filename2 = filename.split('_')[1][-6:]
    account_num = filename.split('_')[1]

    file_paths = {
        "bank": f"{directory}/workF/{filename}.xls",
        "saer": f"{directory}/workF/거래처원장 {filename2}.xls"
    }
    output_file_path = f"{directory}/{account_num}.xlsx"
    output_file_path = output_file_path.replace("workF","resultF")
    output_file_path = util.save_excel_with_seq(output_file_path)

    # 데이터프레임 읽기 및 전처리
    df_bank = pd.read_excel(file_paths["bank"], header=None).iloc[6:]
    df_saer = pd.read_excel(file_paths["saer"], header=None).iloc[7:]
    
    # 초기 저장 (은행 자료, 회계 자료)
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False, header=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False, header=False)
    
    # 5. 특정 시트 읽기    
    df_bank = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["bank"])
    df_saer = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["saer"])

    # 데이터 가공 및 정리
    df_bank = dp.preprocess_bank_data(df_bank)
    df_saer = dp.preprocess_saer_data(df_saer)
    
    # 데이터 병합
    df_combined = dp.combine_df_data(df_bank, df_saer)

    # 피봇 테이블 생성
    df_pivot_out, df_pivot_in = dp.create_pivot_tables(df_combined, SHEET_NAMES["bank"], SHEET_NAMES["saer"])

    # 중간 저장
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False)
        df_combined.to_excel(writer, sheet_name=SHEET_NAMES["combined"], index=False)
        df_pivot_out.to_excel(writer, sheet_name=SHEET_NAMES["pivot_out"])
        df_pivot_in.to_excel(writer, sheet_name=SHEET_NAMES["pivot_in"])

    # 특정 시트 읽기
    df_pivot_out = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["pivot_out"])
    df_pivot_in = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["pivot_in"])

    # 데이터 가공 및 정리
    df_pivot_out = dp.preprocess_pivot_out_data(df_pivot_out)
    df_pivot_in = dp.preprocess_pivot_in_data(df_pivot_in)

    # 피봇데이터 병합
    df_pivot_comb = dp.combine_df_pivot_data(df_pivot_out, df_pivot_in)

    # 최종 저장
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False)
        df_combined.to_excel(writer, sheet_name=SHEET_NAMES["combined"], index=False)
        df_pivot_out.to_excel(writer, sheet_name=SHEET_NAMES["pivot_out"])
        df_pivot_in.to_excel(writer, sheet_name=SHEET_NAMES["pivot_in"])
        df_pivot_comb.to_excel(writer, sheet_name=SHEET_NAMES["error_check"], index=False)