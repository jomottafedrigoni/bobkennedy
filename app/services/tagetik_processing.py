import pandas as pd


def prepare_tagetik(tagetik_df, valid_accounts):

    df = tagetik_df.copy()

    # Tengo solo i conti che esistono nelle transazioni Dynamics
    df = df[
        df["LOCAL_ACCOUNT"].isin(valid_accounts)
    ].copy()

    # Sicurezza numerica
    df["AMOUNT_LC"] = pd.to_numeric(
        df["AMOUNT_LC"],
        errors="coerce"
    )

    # =====================================================
    # AGGREGO PRIMA IL CUMULATO
    # =====================================================

    df = (
        df.groupby(
            [
                "LOCAL_ACCOUNT",
                "LOCAL_ACCOUNT_DESC",
                "TAGETIK BS LINE DESCRIPTION",
                "PERIODO"
            ],
            as_index=False
        )["AMOUNT_LC"]
        .sum()
    )

    # =====================================================
    # ORDINO
    # =====================================================

    df["PERIODO"] = (
        df["PERIODO"]
        .astype(str)
        .str.zfill(2)
    )

    df = df.sort_values(
        [
            "LOCAL_ACCOUNT",
            "TAGETIK BS LINE DESCRIPTION",
            "PERIODO"
        ]
    )

    # =====================================================
    # CUMULATO -> MOVIMENTO
    # =====================================================

    df["MOVIMENTO"] = (
        df.groupby(
            [
                "LOCAL_ACCOUNT",
                "TAGETIK BS LINE DESCRIPTION"
            ]
        )["AMOUNT_LC"]
        .diff()
    )

    first_rows = (
        df.groupby(
            [
                "LOCAL_ACCOUNT",
                "TAGETIK BS LINE DESCRIPTION"
            ]
        )
        .head(1)
        .index
    )

    df.loc[
        first_rows,
        "MOVIMENTO"
    ] = df.loc[
        first_rows,
        "AMOUNT_LC"
    ]

    return df


def apply_bs_translation(
    tagetik_df,
    bs_mapping_df
):

    translated = tagetik_df.merge(
        bs_mapping_df,
        how="left",
        on="TAGETIK BS LINE DESCRIPTION"
    )

    translated = translated[
        translated["Conto"].notna()
    ].copy()

    translated = translated.rename(
        columns={
            "Conto": "Conto Bilancio"
        }
    )

    return translated


def tagetik_pivot(tagetik_df):

    summary = (
        tagetik_df
        .groupby(
            [
                "LOCAL_ACCOUNT",
                "LOCAL_ACCOUNT_DESC",
                "Conto Bilancio",
                "PERIODO"
            ],
            as_index=False
        )["MOVIMENTO"]
        .sum()
    )

    pivot = summary.pivot_table(
        index=[
            "LOCAL_ACCOUNT",
            "LOCAL_ACCOUNT_DESC",
            "Conto Bilancio"
        ],
        columns="PERIODO",
        values="MOVIMENTO",
        aggfunc="sum",
        fill_value=0
    )

    pivot = pivot.reset_index()

    return pivot