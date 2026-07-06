import pandas as pd


def check_mapping_duplicates(mapping_df):

    duplicates = mapping_df[
        mapping_df.duplicated(
            subset=[
                "Conto principale",
                "Cost Center"
            ],
            keep=False
        )
    ]

    return duplicates


def enrich_transactions(transactions_df, mapping_df):

    merged = transactions_df.merge(
        mapping_df,
        how="left",
        on=[
            "Conto principale",
            "Cost Center"
        ]
    )

    return merged


def get_unmapped_transactions(df):

    return df[
        df["Conto Bilancio"].isna()
    ]

def build_dynamics_pivot(df):

    work_df = df[
        df["Conto Bilancio"].notna()
    ].copy()

    work_df["Periodo"] = (
        pd.to_datetime(work_df["Data"])
        .dt.month
        .astype(str)
        .str.zfill(2)
    )

    summary = (
        work_df
        .groupby(
            [
                "Conto principale",
                "Nome conto",
                "Conto Bilancio",
                "Periodo"
            ]
        )["Importo"]
        .sum()
        .reset_index()
    )

    pivot = summary.pivot_table(
        index=[
            "Conto principale",
            "Nome conto",
            "Conto Bilancio"
        ],
        columns="Periodo",
        values="Importo",
        aggfunc="sum",
        fill_value=0
    )

    pivot = pivot.reset_index()

    return pivot