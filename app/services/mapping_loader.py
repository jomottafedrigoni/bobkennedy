from pathlib import Path
import pandas as pd

def load_mapping(root_path: Path) -> pd.DataFrame:

    file_path = (
        root_path
        / "data"
        / "Transazioni"
        / "RITR Conti di pertinenza.xlsx"
    )

    df = pd.read_excel(
        file_path,
        sheet_name="Lista Conti",
        header=1
    )

    df = df.rename(
        columns={
            "DESC_ITA_MAN": "Conto Bilancio",
            "CostCenter": "Cost Center",
            "MainAccount": "Conto principale"
        }
    )

    cols = [
        "Conto principale",
        "Cost Center",
        "Conto Bilancio",
        "Scarico"
    ]

    df = df[cols].copy()

    return df