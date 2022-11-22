import pandas as pd
import xlwings as xw

from datetime import datetime, date
import os
from pathlib import Path

import functions as fn


def main():
    collecting_data()

def collecting_data():
    df = fn.get_caller_df().copy()

    #gather EMS-information and merges with df
    df_ems = fn.get_csv_data('ems').copy()
    df_ems = df_ems.rename(columns={'UNNO':'UNNR'})
    df = df.merge(df_ems, how='left', on='UNNR').copy()

    df = fn.regex_no_extra_whitespace(df).copy()
    df_ems = fn.regex_no_extra_whitespace(df_ems).copy()

    df.dropna(subset=['IMDG'], inplace=True)

    df_dg = pd.DataFrame(columns=[
        'MLO', 'Reference', 'TOL', 'SIZE', 'CONTAINER#', 'IMO class', 'UN no', 'IMO Name',
        'Package Group', 'MP', 'FP (°C)', 'NO. OF PK', 'Packages ', 'Gross weight ( kg )',
        'Net weight ( kg )', 'EMS', 'POD', 'Acceptance ref'
        ])

    df_dg.loc[:, 'MLO'] = df['MLO']
    df_dg.loc[:, 'Reference'] = df['BOOKING NUMBER']
    df_dg.loc[:, 'TOL'] = df['TOL']
    df_dg.loc[:, 'SIZE'] = df['ISO TYPE']
    df_dg.loc[:, 'CONTAINER#'] = df['CONTAINER']
    df_dg.loc[:, 'IMO class'] = df['IMDG']
    df_dg.loc[:, 'UN no'] = df['UNNR']
    df_dg.loc[:, 'IMO Name'] = ""
    df_dg.loc[:, 'Package Group'] = ""
    df_dg.loc[:, 'MP'] = ""
    df_dg.loc[:, 'FP (°C)'] = ""
    df_dg.loc[:, 'NO. OF PK'] = ""
    df_dg.loc[:, 'Packages'] = ""
    df_dg.loc[:, 'Gross weight ( kg )'] = ""
    df_dg.loc[:, 'Net weight ( kg )'] = ""
    df_dg.loc[:, 'EMS'] = df['EMS']
    df_dg.loc[:, 'POD'] = df['FINAL POD']
    df_dg.loc[:, 'Acceptance ref'] = df['CHEM REF']
   
    return finish(df_dg)

def finish(df: pd.DataFrame):

    vessel = fn.get_caller_df.vessel
    voyage = fn.get_caller_df.voyage
    pol = fn.get_caller_df.pol
    today = date.today().strftime("%Y-%m-%d")
    len_df = len(df)-1

    #find user then create path to local python_templates
    wb_caller_path = xw.Book.caller().fullname
    p = Path(wb_caller_path)
    user = p.parts[2]
    head = r'C:\Users'
    tail = r'Documents\python_templates\template-dcm.xlsx'
    dcm_template = os.path.join(head, user, tail)

    # get desktop name
    desktop_swe = r"OneDrive - BOLLORE\Skrivbordet"
    desktop_eng = r"OneDrive - BOLLORE\Desktop"
    desktop_path_swe = os.path.join(head, user, desktop_swe)
    desktop_path_eng = os.path.join(head, user, desktop_eng)
    desktop_path_swe_without_onedrive = desktop_path = os.path.join(head, user, 'Skrivbord')
    desktop_path_eng_without_onedrive = desktop_path = os.path.join(head, user, 'Desktop')
    desktop_path = ""

    # run if statement if desktop in swe or eng
    if os.path.exists(desktop_path_swe_without_onedrive):
        desktop_path = desktop_path_swe_without_onedrive

    elif os.path.exists(desktop_path_eng_without_onedrive):
        desktop_path = desktop_path_eng_without_onedrive

    elif os.path.exists(desktop_path_swe):
        desktop_path = desktop_path_swe

    else:
        desktop_path = desktop_path_eng

    time_str = datetime.now().strftime("%y%m%d")
    dgm_file_name = "DG_" + vessel + "_" + str(voyage) + "_" + pol + "_" + time_str + ".xlsx"
    dgm_full_path = os.path.join(desktop_path, dgm_file_name)

    with xw.App(visible=False) as app:
        wb = app.books.open(dcm_template)
        wb.save(dgm_full_path)

        dcm_sheet = wb.sheets['DCM']
        dcm_sheet.range('C11').value = vessel
        dcm_sheet.range('F11').value = voyage
        dcm_sheet.range('I11').value = pol
        dcm_sheet.range('F18').value = today
        dcm_sheet.range((14, 1), (14 + len_df, 19)).insert('down')
        dcm_sheet.range('B14').options(pd.DataFrame, index=False, header=False).value = df.copy()

        wb.save()
        wb.close()

if __name__ == '__main__':
    file_path = fn.get_mock_caller('0109_Bokningsblad.xlsb')
    xw.Book(file_path).set_mock_caller()
    collecting_data()
    