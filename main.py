import pandas as pd
from glob import glob
import os

if __name__ == "__main__":
    
    # MONTH, YEAR = 6, 2024
    MONTH = int(input("Enter the month: "))
    YEAR = int(input("Enter the year: "))
    month = str(MONTH).zfill(2)
    year = str(YEAR)

    outbound_file_path = f"M:/CPP-Data/Sutherland RPA/BD IS Printing"
    criteria = f"/{year}/{year}{month}*_*Outbound*.xlsx"

    this_month_outbound_files = glob(f"{outbound_file_path}/{criteria}")
    curr_month = pd.concat([pd.read_excel(file) for file in this_month_outbound_files], ignore_index=True)
    
    if MONTH == 1:
        last_month = str(12).zfill(2)
        last_year = str(YEAR-1)
    else:
        last_month = str(MONTH-1).zfill(2)
        last_year = year


    criteria = f"/{last_year}/{last_year}{last_month}*_*Outbound*.xlsx"
    last_month_outbound_files = glob(f"{outbound_file_path}/{criteria}")
    last_month = pd.concat([pd.read_excel(file) for file in last_month_outbound_files], ignore_index=True)
    
    df = pd.concat([curr_month, last_month], ignore_index=True)
    df = df.drop_duplicates()
    
    keep_cols = ['POLICYID','Category', 'BotName','RetrievalStatus','RetrievalDescription','CreatedDate','LastModifiedDate']
    df = df[keep_cols]
    
    # keep rows where the createddate is within the same month as MONTH
    df['Month'] = df['CreatedDate'].dt.month
    df['Year']  = df['CreatedDate'].dt.year

    # keep rows where df['Year'] == curr_year and df['Month'] == curr_month
    df = df[(df['Year'] == YEAR) & (df['Month'] == MONTH)]
    
    pivot = df.pivot_table(index='BotName', columns='RetrievalStatus', values='POLICYID', aggfunc='count')
    
    os.makedirs('./results', exist_ok=True)
    
    with pd.ExcelWriter(f'./results/{year} {month} - Itemized Statement Invoicing.xlsx') as writer:
        df.to_excel(writer, sheet_name='Itemized Statement', index=False)
        pivot.to_excel(writer, sheet_name='Pivot Table')