import pandas as pd
import openpyxl as opxl
import numpy as np
import os
import json
from datetime import datetime as dt


def fetchExcel(path: str, wbOld: str, wbNew: str):
    dfOld = pd.read_excel(
        os.path.join(path, wbOld)
    )
    dfNew = pd.read_excel(
        os.path.join(path, wbNew)
    )

    dp = {
        'dfOld': dfOld,
        'dfNew': dfNew
    }
    return dp


def comparison(dp: dict):
    dfOld = dp['dfOld']
    dfNew = dp['dfNew']

    added = dfNew[~dfNew.index.isin(dfOld.index)]
    removed = dfOld[~dfOld.index.isin(dfNew.index)]
    common = dfNew[dfNew.index.isin(dfOld.index)]
    changes = (common != dfOld.loc[common.index])

    print(f"Newly added Data:\n{added}")
    print(f"Removed Data:\n{removed}")
    print(f"Common Data:\n{common}")
    print(f"Changes in Data:\n{changes}")


def main():
    with open('variables.json', 'r') as j:
        data = json.load(j)
    path = data['path']
    wbOld = data['oldFile']
    wbNew = data['newFile']
    dp = fetchExcel(path, wbOld, wbNew)
    comparison(dp)


if __name__ == "__main__":
    main()
