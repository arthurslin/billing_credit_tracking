import pandas as pd
import datetime
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
            if df.columns[0] == "BILLING REPORT DETAIL":
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)
        to_return.append(df)
    return to_return

def get_credits(data):
    billing_report = data[0]
    credits = billing_report.drop(
        billing_report[billing_report['Part Number'] != '778752'].index, inplace=False)
    
    po_matches = credits["PO Number"].tolist() 
    training_orders = billing_report[billing_report["PO Number"].isin(po_matches)] #Grab all matched items
    training_orders["Credits"] = training_orders["Actual Ship Date"] + pd.DateOffset(years=1)
    items_to_extract = ["Order Number","Region","BU","Product Family","Customer Name","PO Number","Part Number","Part Description","Sales Person","Order Quantity","Line Total (Local) ","Actual Ship Date","Credits"]
    clean_items = training_orders[items_to_extract]
    clean_items["Actual Ship Date"] = pd.to_datetime(clean_items["Actual Ship Date"]).dt.date
    clean_items["Credits"] = pd.to_datetime(clean_items["Credits"]).dt.date
    clean_items["Today"] = datetime.datetime.today().date()
    clean_items["Expired?"] = clean_items["Today"] > clean_items["Credits"]

    clean_items.to_excel("order.xlsx",index=False)


dirs = ["billing_report"]

credit_data = load_data(dirs)

get_credits(credit_data)