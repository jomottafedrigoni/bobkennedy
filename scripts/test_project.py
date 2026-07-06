from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

print(f"Root progetto: {ROOT}")

transazioni = ROOT / "data" / "Transazioni"
industriale = ROOT / "data" / "Dati Industriali"

print("Transazioni:", transazioni.exists())
print("Dati Industriali:", industriale.exists())