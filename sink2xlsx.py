import pandas as pd
from sqlalchemy import create_engine

engine = create_engine('sqlite:///result/operations.db')
frame = pd.read_sql('qcc_firm_detail', engine)
print(frame)
frame.to_excel('./result/result.xlsx')
