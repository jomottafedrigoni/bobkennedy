from pathlib import Path
import pandas as pd


def load_bs_mapping(root_path: Path):

    file_path = (
        root_path
        / "data"
        / "Transazioni"
        / "ENG ITA BS Accounts.xlsx"
    )

    df = pd.read_excel(
        file_path,
        header=1
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df