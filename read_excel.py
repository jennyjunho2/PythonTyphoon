import pandas as pd
import os
from datetime import datetime

def read_problem_from_excel(excel_file, year, test_paper_name):
    if year == 1:
        df = pd.read_excel(excel_file, sheet_name = '고1 '+ str(datetime.today().year))
    elif year == 2:
        df = pd.read_excel(excel_file, sheet_name = '고2 '+ str(datetime.today().year))
    else:
        print("학년이 맞는지 확인해 주세요.")
    print(df.loc[:, '대단원', '소단원', '난이도', '번호', test_paper_name])


sample_excel = r"C:\Users\이준호\PycharmProjects\태풍\내신주문서.xlsx"
read_problem_from_excel(sample_excel, 1, '신목 최종마무리2 (210503)')