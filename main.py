import pandas as pd
import openpyxl as opxl
import numpy as np
import os
import json
from datetime import datetime as dt


def fetchJSON():
    # Fetch data from JSON file
    with open('variables.json', 'r') as j:
        var = json.load(j)

    return var


def fetchExcel(var: dict):
    dfOld = pd.read_excel(
        os.path.join(var['path'], var['oldFile']),
        sheet_name=var['oldSheet']
    )
    print(f"Old Data Frame:\n{dfOld}")

    dfNew = pd.read_excel(
        os.path.join(var['path'], var['newFile']),
        sheet_name=var['newSheet']
    )
    print(f"New Data Frame:\n{dfNew}")

    dp = {
        'dfOld': dfOld,
        'dfNew': dfNew
    }
    return dp


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


def main():
    # Fetch JSON file for variables
    var = fetchJSON()

    # Fetch Excel files to compare
    dp = fetchExcel(var)

    # Run comparison based on index
    print("\nIndex Based Comparison")
    comparison(dp, None)
    print("_"*100)

    # Run comparison based on specific column
    print("\nColumn Based Comparison")
    comparison(dp, var['col'])
    print("_"*100)


if __name__ == "__main__":
    main()
