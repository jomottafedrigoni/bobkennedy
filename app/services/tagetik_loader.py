from pathlib import Path
import pandas as pd


def load_tagetik(root_path: Path):

    file_path = (
        root_path
        / "data"
        / "Transazioni"
        / "tagetik.xlsx"
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

    df["LOCAL_ACCOUNT"] = (
        df["LOCAL_ACCOUNT"]
        .astype(str)
        .str.strip()
    )

    df["PERIODO"] = (
        df["PERIODO"]
        .astype(str)
        .str.zfill(2)
    )

    return df