import pandas as pd
from openpyxl import load_workbook as loadWB
import os
import json
from datetime import datetime as dt
from formatz import tableFormat


def fetchJSON():
    # Fetch data from JSON file
    with open('variables.json', 'r') as j:
        var = json.load(j)

    return var


def fetchExcel(var: dict):
    dfOld = pd.read_excel(
        os.path.join(var['src'], var['oldFile']),
        sheet_name=var['oldSheet']
    )
    print(f"\nOld Data Frame:\n{dfOld}")

    dfNew = pd.read_excel(
        os.path.join(var['src'], var['newFile']),
        sheet_name=var['newSheet']
    )
    print(f"\nNew Data Frame:\n{dfNew}")

    return {
        'dfOld': dfOld,
        'dfNew': dfNew
    }


def comparison(dp: dict, col: str):
    dfOld = dp['dfOld'].copy()
    dfNew = dp['dfNew'].copy()

    if col:
        # Set col as index for comparison
        dfOld.set_index(col, inplace=True)
        dfNew.set_index(col, inplace=True)

        # Align columns
        commonColz = dfOld.columns.intersection(dfNew.columns)
        dfOld = dfOld[commonColz]
        dfNew = dfNew[commonColz]

    # Identify differences
    added = dfNew[~dfNew.index.isin(dfOld.index)]
    removed = dfOld[~dfOld.index.isin(dfNew.index)]
    common = dfNew[dfNew.index.isin(dfOld.index)]
    changes = (common != dfOld.loc[common.index])

    # Output
    print(f"\n🟢 Newly Added Rows:\n{added.reset_index()}")
    print(f"\n🔴 Removed Rows:\n{removed.reset_index()}")
    print(f"\n🔁 Common Rows (new):\n{common.reset_index()}")
    print(f"\n🟡 Changes Detected (True where value changed):\n{changes}")

    # Optionally return these
    return {
        'added': added.reset_index(),
        'removed': removed.reset_index(),
        'common': common.reset_index(),
        'changes': changes.reset_index()
    }


def toExcel(data: dict, var: dict, ty: str):
    if ty == 'index':
        wbN = os.path.join(var['wrk'], var['outFile1'])
    if ty == 'col':
        wbN = os.path.join(var['wrk'], var['outFile2'])

    with pd.ExcelWriter(wbN) as xlwriter:
        for key, value in data.items():
            if isinstance(value, pd.DataFrame):
                value.to_excel(xlwriter, sheet_name=key, index=False)

    wb = loadWB(wbN)
    tableFormat(wb)
    wb.save(wbN)


def main():
    # Fetch JSON file for variables
    var = fetchJSON()

    # Fetch Excel files to compare
    print("Excel Files for comparison")
    dp = fetchExcel(var)
    print("_"*100)

    # Run comparison based on index
    print("Index Based Comparison")
    data = comparison(dp, None)
    toExcel(data, var, 'index')
    print("_"*100)

    # Run comparison based on specific column
    print("Column Based Comparison")
    data = comparison(dp, var['col'])
    toExcel(data, var, 'col')
    print("_"*100)


if __name__ == "__main__":
    main()
