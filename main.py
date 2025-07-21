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
    print(f"Old Data Frame:\n{dfOld}")

    dfNew = pd.read_excel(
        os.path.join(path, wbNew)
    )
    print(f"New Data Frame:\n{dfNew}")

    dp = {
        'dfOld': dfOld,
        'dfNew': dfNew
    }
    return dp


def unorderedComparison(dp: dict):
    dfOld = dp['dfOld'].copy()
    dfNew = dp['dfNew'].copy()

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


def orderedComparison(dp: dict, col: str):
    dfOld = dp['dfOld'].copy()
    dfNew = dp['dfNew'].copy()

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
    with open('variables.json', 'r') as j:
        data = json.load(j)
    path = data['path']
    wbOld = data['oldFile']
    wbNew = data['newFile']
    dp = fetchExcel(path, wbOld, wbNew)
    unorderedComparison(dp)


if __name__ == "__main__":
    main()
