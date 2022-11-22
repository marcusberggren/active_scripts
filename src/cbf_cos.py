import pandas as pd
from pandas.api.types import CategoricalDtype
import xlwings as xw

from datetime import datetime
from pathlib import Path
import os

from functions import get_path


def main():

    wb = xw.Book.caller()
    sheet = wb.sheets("INFO")

    vessel = sheet.range("A2").value
    alt_voy = sheet.range("F2").value
    pol = sheet.range("C2").value

    if alt_voy is None:
        alt_voy = "TBA"

    cell_range = sheet.range("A4").expand()  # dynamisk range
    df = (
        sheet.range(cell_range).options(pd.DataFrame, index=False, header=True).value
    )  # dynamisk range

    df = df[["TOL", "ISO TYPE", "LOAD STATUS", "NET WEIGHT", "VGM"]].copy()

    # När NET WEIGHT är mindre än 100 men större än 0 så multiplicera med 1000
    df.loc[(df["NET WEIGHT"] < 100) & (df["NET WEIGHT"] != 0), "NET WEIGHT"] *= 1000

    # Skapar ny kolumn med maxvärde av NET WEIGHT & VGM
    df.loc[:, "MAX WEIGHT"] = df[["NET WEIGHT", "VGM"]].max(axis=1) // 1000

    df = df[["TOL", "ISO TYPE", "LOAD STATUS", "MAX WEIGHT"]].copy()

    df.loc[df["LOAD STATUS"] != "MT", "LOAD STATUS"] = "LA"

    df.insert(4, "WEIGHT_TYPE", "")

    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[0] == "2")
        & (df["MAX WEIGHT"] >= 20),
        "WEIGHT_TYPE",
    ] = "VH20"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[0] == "2")
        & (df["MAX WEIGHT"] >= 15)
        & (df["MAX WEIGHT"] < 20),
        "WEIGHT_TYPE",
    ] = "H20"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[0] == "2")
        & (df["MAX WEIGHT"] >= 10)
        & (df["MAX WEIGHT"] < 15),
        "WEIGHT_TYPE",
    ] = "M20"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[0] == "2")
        & (df["MAX WEIGHT"] < 10),
        "WEIGHT_TYPE",
    ] = "L20"

    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "42")
        & (df["MAX WEIGHT"] >= 25),
        "WEIGHT_TYPE",
    ] = "VH42"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "42")
        & (df["MAX WEIGHT"] >= 20)
        & (df["MAX WEIGHT"] < 25),
        "WEIGHT_TYPE",
    ] = "H42"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "42")
        & (df["MAX WEIGHT"] >= 15)
        & (df["MAX WEIGHT"] < 20),
        "WEIGHT_TYPE",
    ] = "M42"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "42")
        & (df["MAX WEIGHT"] < 15),
        "WEIGHT_TYPE",
    ] = "L42"

    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "45")
        & (df["MAX WEIGHT"] >= 25),
        "WEIGHT_TYPE",
    ] = "VH45"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "45")
        & (df["MAX WEIGHT"] >= 20)
        & (df["MAX WEIGHT"] < 25),
        "WEIGHT_TYPE",
    ] = "H45"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "45")
        & (df["MAX WEIGHT"] >= 15)
        & (df["MAX WEIGHT"] < 20),
        "WEIGHT_TYPE",
    ] = "M45"
    df.loc[
        (df["LOAD STATUS"] == "LA")
        & (df["ISO TYPE"].str[:2] == "45")
        & (df["MAX WEIGHT"] < 15),
        "WEIGHT_TYPE",
    ] = "L45"

    df.loc[
        (df["LOAD STATUS"] == "MT") & (df["ISO TYPE"].str[0] == "2"), "WEIGHT_TYPE"
    ] = "MT20"
    df.loc[
        (df["LOAD STATUS"] == "MT") & (df["ISO TYPE"].str[:2] == "42"), "WEIGHT_TYPE"
    ] = "MT42"
    df.loc[
        (df["LOAD STATUS"] == "MT") & (df["ISO TYPE"].str[:2] == "45"), "WEIGHT_TYPE"
    ] = "MT45"

    lista_iso_type = [
        "VH20",
        "VH42",
        "VH45",
        "H20",
        "H42",
        "H45",
        "M20",
        "M42",
        "M45",
        "L20",
        "L42",
        "L45",
        "MT20",
        "MT42",
        "MT45",
    ]

    cat_size_order = CategoricalDtype(lista_iso_type, ordered=True)

    df["WEIGHT_TYPE"] = df["WEIGHT_TYPE"].astype(cat_size_order)

    df.sort_values(["WEIGHT_TYPE", "TOL"], ascending=(True, False))

    df = (
        df.groupby(["TOL", "WEIGHT_TYPE"]).size().iteritems()
    )  # groupby och size för att skapa hanterlig strukturerad data. Kan iterera över datan nedan mha iteritems

    dict1 = {}

    for (
        TOL,
        iso_type,
    ), antal in df:  # tar fram nestlad info (TOL, iso_type) och antal från ovan
        if TOL not in dict1:
            dict1[TOL] = {}
        dict1[TOL].update({iso_type: antal})

    TOL = ""
    index = 0

    lista_weight_type = []
    create_nested_list = []
    lista_TOL = []
    lista_unika_TOLer = []

    for TOL in dict1:
        lista_unika_TOLer.append(TOL)

        for index, weight_type in enumerate(dict1[TOL].values()):

            if index % 3 == 2:
                lista_weight_type.append(weight_type)
                create_nested_list.append(lista_weight_type)
                lista_weight_type = []
            else:
                lista_weight_type.append(weight_type)
        lista_TOL.append(create_nested_list)
        create_nested_list = []

    df_new = pd.DataFrame(lista_TOL)

    def get_TOL(lista_TOLer):
        i = 0
        for i, TOL in enumerate(lista_TOLer):

            if TOL == "NLEDE":
                lista_TOLer[i] = "RTM-DDE"
            elif TOL == "NLEMX":
                lista_TOLer[i] = "RTM-EMX"
            elif TOL == "NLRWG":
                lista_TOLer[i] = "RTM-RWG"
            elif TOL == "DECTB":
                lista_TOLer[i] = "HAM-CTB"
            elif TOL == "DECTA":
                lista_TOLer[i] = "HAM-CTA"
            elif TOL == "DETCT":
                lista_TOLer[i] = "HAM-CTT"

        return lista_TOLer

    # find user then create path to local python_templates
    wb_caller_path = xw.Book.caller().fullname
    p = Path(wb_caller_path)
    user = p.parts[2]
    head = r"C:\Users"
    tail = r"Documents\python_templates\template-cbf-cos.xlsx"
    cbf_template = os.path.join(head, user, tail)

    # get desktop name
    desktop_swe = r"OneDrive - BOLLORE\Skrivbordet"
    desktop_eng = r"OneDrive - BOLLORE\Desktop"
    desktop_path_swe = os.path.join(head, user, desktop_swe)
    desktop_path_eng = os.path.join(head, user, desktop_eng)
    desktop_path_swe_without_onedrive = desktop_path = os.path.join(
        head, user, "Skrivbord"
    )
    desktop_path_eng_without_onedrive = desktop_path = os.path.join(
        head, user, "Desktop"
    )
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

    with xw.App(visible=False) as app:
        wb = app.books.open(cbf_template)
        wb_caller = xw.Book.caller()
        wb_caller_name = wb_caller.fullname
        time_str = datetime.now().strftime("%y%m%d")

        filename = (
            "CBF_" + vessel + "_" + str(alt_voy) + "_" + pol + "_" + time_str + ".xlsx"
        )
        dir_path = os.path.join(desktop_path, filename)

        wb.save(dir_path)

        ws = wb.sheets["CBF TTL"]
        ws.range("B3").value = vessel
        ws.range("B4").value = alt_voy
        ws.range("I3").value = pol

        TOL, cell, i = "", 8, 0
        for i, TOL in enumerate(get_TOL(lista_unika_TOLer)):

            ws.range("A" + str(cell + 1)).options(index=False, header=False).value = TOL

            ws.range("C" + str(cell)).options(index=False, header=False).value = df_new[
                0
            ][i]
            ws.range("C" + str(cell + 1)).options(
                index=False, header=False
            ).value = df_new[1][i]
            ws.range("C" + str(cell + 2)).options(
                index=False, header=False
            ).value = df_new[2][i]
            ws.range("C" + str(cell + 3)).options(
                index=False, header=False
            ).value = df_new[3][i]
            ws.range("C" + str(cell + 4)).options(
                index=False, header=False
            ).value = df_new[4][i]
            cell += 7

        wb.save()
        wb.close()
