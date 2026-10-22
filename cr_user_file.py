"""Write a user file with CR in `costCenter` for Microsoft's chargeback tools (CreditUsage, FinOps Cowork).

Usage: python cr_user_file.py <data_dir> <YYYY-MM> [out.csv]
Users come from that month's admin-center export; CR uses the same rules as build_report.py.
"""
import pathlib
import sys

import pandas as pd

from build_report import assign_cr, load_cr, load_exports


def user_file(d: pathlib.Path, month: str) -> pd.DataFrame:
    users = load_exports(d).query("month == @month")[["month", "upn", "name"]].drop_duplicates("upn")
    if users.empty:
        sys.exit(f"No M365AIConsumptionByUser export for {month} in {d}")
    users = users.reset_index(drop=True)
    users["CR"] = assign_cr(users, load_cr(d)).replace("UNMAPPED", "NO-CR")
    return pd.DataFrame({"userPrincipalName": users["upn"], "displayName": users["name"],
                         "department": users["CR"], "costCenter": users["CR"]})


if __name__ == "__main__":
    d, month = pathlib.Path(sys.argv[1]), sys.argv[2]
    out = pathlib.Path(sys.argv[3]) if len(sys.argv) > 3 else d / f"entra_with_cr_{month}.csv"
    df = user_file(d, month)
    df.to_csv(out, index=False)
    print(f"{out}: {len(df)} users, {(df['costCenter'] == 'NO-CR').sum()} without CR")
