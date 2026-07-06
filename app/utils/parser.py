import pandas as pd


def parse_account_display_value(value):

    if pd.isna(value):
        return pd.Series(
            [None, None, None, None]
        )

    parts = str(value).split("-")

    conto = parts[0] if len(parts) > 0 else None
    cost_center = parts[1] if len(parts) > 1 else None
    plant = parts[2] if len(parts) > 2 else None

    fornitore = None

    for part in reversed(parts):
        if part.strip():
            fornitore = part.strip()
            break

    return pd.Series(
        [
            conto,
            cost_center,
            plant,
            fornitore
        ]
    )