import pandas as pd
import glob
import os

def load_data(directories):
    to_return = []
    for i in directories:
        paths = glob.glob(os.path.join(i, "*xlsx"))
        if not paths:
            raise FileNotFoundError(i, "File not found")
        for path in paths:
            xl = pd.ExcelFile(path)
            df = pd.read_excel(path, sheet_name=xl.sheet_names[0])
        to_return.append(df)
    return to_return

def get_credits(data):
    pass


dirs = []

credit_data = load_data()