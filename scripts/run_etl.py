from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent

sys.path.append(str(ROOT))

from app.services.mapping_loader import load_mapping
from app.services.transaction_loader import load_transactions

from app.services.reconciliation import (
    enrich_transactions,
    build_dynamics_pivot
)

from app.services.reconciliation_builder import (
    build_reconciliation
)

from app.services.tagetik_loader import load_tagetik

from app.services.tagetik_processing import (
    prepare_tagetik,
    apply_bs_translation,
    tagetik_pivot
)

from app.services.bs_mapping_loader import (
    load_bs_mapping
)

from app.utils.parser import (
    parse_account_display_value
)


# ====================================================
# CACHE
# ====================================================

CACHE_DIR = ROOT / "data" / "cache"

CACHE_DIR.mkdir(exist_ok=True)

# ====================================================
# MAPPING
# ====================================================

print("Caricamento mapping...")

mapping_df = load_mapping(ROOT)

# ====================================================
# TRANSAZIONI
# ====================================================

print("Caricamento transazioni...")

transactions_df = load_transactions(ROOT)

print("Parsing account display value...")

transactions_df[
    [
        "Conto Estratto",
        "Cost Center",
        "Plant",
        "Fornitore"
    ]
] = transactions_df[
    "Account display value"
].apply(
    parse_account_display_value
)

print("Applicazione mapping...")

enriched_df = enrich_transactions(
    transactions_df,
    mapping_df
)

test = enriched_df[
    enriched_df["Descrizione"]
    .str.contains("mensa", case=False, na=False)
]

print(
    test[
        [
            "Origine",
            "Numero giornale di registrazione",
            "Conto principale",
            "Cost Center",
            "Descrizione",
            "Importo"
        ]
    ]
)

# ====================================================
# DYNAMICS PIVOT
# ====================================================

valid_accounts = (
    enriched_df[
        enriched_df["Conto Bilancio"].notna()
    ]["Conto principale"]
    .astype(str)
    .unique()
)

print("Creazione pivot Dynamics...")

dynamics_pivot = build_dynamics_pivot(
    enriched_df
)

# ====================================================
# TAGETIK
# ====================================================

print("Caricamento Tagetik...")

tagetik_df = load_tagetik(ROOT)

tagetik_df = prepare_tagetik(
    tagetik_df,
    valid_accounts
)

# ====================================================
# TRADUZIONE ENG -> ITA
# ====================================================

print("Caricamento mapping BS...")

bs_mapping_df = load_bs_mapping(ROOT)

print("Applicazione traduzione BS...")

tagetik_df = apply_bs_translation(
    tagetik_df,
    bs_mapping_df
)

print(
    f"Righe Tagetik dopo traduzione: {len(tagetik_df):,}"
)

# ====================================================
# PIVOT TAGETIK
# ====================================================

print("Creazione pivot Tagetik...")

tagetik_pivot_df = tagetik_pivot(
    tagetik_df
)

print("Creazione riconciliazione...")

reconciliation_df = build_reconciliation(
    dynamics_pivot,
    tagetik_pivot_df
)


# ====================================================
# EXPORT CACHE
# ====================================================

print("Salvataggio cache...")

enriched_df.to_parquet(
    CACHE_DIR / "enriched_transactions.parquet",
    index=False
)

dynamics_pivot.to_parquet(
    CACHE_DIR / "dynamics_pivot.parquet",
    index=False
)

tagetik_df.to_parquet(
    CACHE_DIR / "tagetik_clean.parquet",
    index=False
)

tagetik_pivot_df.to_parquet(
    CACHE_DIR / "tagetik_pivot.parquet",
    index=False
)

reconciliation_df.to_parquet(
    CACHE_DIR / "reconciliation.parquet",
    index=False
)

print("ETL completato")