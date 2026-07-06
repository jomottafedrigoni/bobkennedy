import pandas as pd

def build_reconciliation(
    dynamics_pivot,
    tagetik_pivot
):

    tagetik_compare = tagetik_pivot.rename(
        columns={
            "LOCAL_ACCOUNT": "Conto principale",
            "LOCAL_ACCOUNT_DESC": "Nome conto Tagetik"
        }
    )

    reconciliation = dynamics_pivot.merge(
        tagetik_compare,
        how="outer",
        on=[
            "Conto principale",
            "Conto Bilancio"
        ],
        suffixes=(
            "_DYN",
            "_TAG"
        )
    )

    months = [
        "01", "02", "03", "04",
        "05", "06", "07", "08",
        "09", "10", "11", "12"
    ]

    for month in months:

        dyn_col = f"{month}_DYN"
        tag_col = f"{month}_TAG"

        if dyn_col not in reconciliation.columns:
            reconciliation[dyn_col] = 0

        if tag_col not in reconciliation.columns:
            reconciliation[tag_col] = 0

        reconciliation[dyn_col] = (
            reconciliation[dyn_col]
            .fillna(0)
        )

        reconciliation[tag_col] = (
            reconciliation[tag_col]
            .fillna(0)
        )

        reconciliation[f"{month}_DELTA"] = (
            reconciliation[dyn_col]
            -
            reconciliation[tag_col]
        ).round(2)

    return reconciliation