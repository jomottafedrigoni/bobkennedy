from pathlib import Path
import pandas as pd


def load_single_file(file_path: Path, origine: str) -> pd.DataFrame:

    df = pd.read_excel(
        file_path
    )

    print(f"\nFILE: {file_path.name}")
    print(df.columns.tolist())

    df.columns = df.columns.str.strip()


    # Conversione importo
    df["Importo"] = (
        df["Importo"]
        .astype(str)
        .str.replace(",", "", regex=False)
    )

    df["Importo"] = pd.to_numeric(
        df["Importo"],
        errors="coerce"
    )

    df["Origine"] = origine

    columns_to_keep = [
    "Numero giornale di registrazione",
    "Giustificativo",
    "Data",
    "Data documento",
    "Anno chiuso",
    "Conto principale",
    "Account display value",
    "Nome conto",
    "Descrizione",
    "Valuta",
    "Importo",
    "Origine"
    ]

    df = df[columns_to_keep]

    return df


def load_transactions(root_path: Path) -> pd.DataFrame:

    trans_folder = root_path / "data" / "Transazioni"

    files = {
        "OVDs.xlsx": "OVDs",
        "VICs.xlsx": "VICs",
        "CF Plant.xlsx": "CF Plant"
    }

    frames = []

    for file_name, origine in files.items():

        file_path = trans_folder / file_name

        frames.append(
            load_single_file(file_path, origine)
        )

    all_transactions = pd.concat(
        frames,
        ignore_index=True
    )

    print("OVDs:", len(frames[0]))
    print("VICs:", len(frames[1]))
    print("CF Plant:", len(frames[2]))

    print("Totale dopo concat:", len(all_transactions))

    return all_transactions

    return pd.concat(frames, ignore_index=True)