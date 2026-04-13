"""
PROGETTO 1 – Analisi Esplorativa Dataset Vendite E-commerce
============================================================
Dataset: Online Retail II (UCI Machine Learning Repository)
         https://archive.ics.uci.edu/dataset/502/online+retail+ii

Come usare:
1. Scarica il file "online_retail_II.xlsx" dal link sopra
2. Mettilo nella stessa cartella di questo script
3. Esegui: python analisi_vendite.py

Output: grafici salvati come PNG + stampa dei principali insight
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import warnings
warnings.filterwarnings("ignore")

# ── Configurazione stile ────────────────────────────────────────────────────
sns.set_theme(style="whitegrid", palette="Blues_d")
plt.rcParams["figure.dpi"] = 120
plt.rcParams["font.family"] = "DejaVu Sans"

# ── 1. CARICAMENTO E PULIZIA ────────────────────────────────────────────────
print("=" * 60)
print("ANALISI ESPLORATIVA – VENDITE E-COMMERCE")
print("=" * 60)

print("\n[1/5] Caricamento dati...")
df = pd.read_excel("online_retail_II.xlsx", sheet_name="Online Retail")

print(f"      Righe caricate: {len(df):,}")
print(f"      Colonne: {list(df.columns)}")

print("\n[2/5] Pulizia dati...")
df_clean = df.copy()

# Rimuovi righe senza CustomerID e con quantità o prezzo negativi
df_clean = df_clean.dropna(subset=["CustomerID"])
df_clean = df_clean[df_clean["Quantity"] > 0]
df_clean = df_clean[df_clean["UnitPrice"] > 0]

# Rimuovi cancellazioni (codici che iniziano con 'C')
df_clean = df_clean[~df_clean["InvoiceNo"].astype(str).str.startswith("C")]

# Colonna revenue
df_clean["Revenue"] = df_clean["Quantity"] * df_clean["UnitPrice"]

# Colonne temporali
df_clean["InvoiceDate"] = pd.to_datetime(df_clean["InvoiceDate"])
df_clean["Month"] = df_clean["InvoiceDate"].dt.to_period("M")
df_clean["Hour"] = df_clean["InvoiceDate"].dt.hour
df_clean["DayOfWeek"] = df_clean["InvoiceDate"].dt.day_name()

print(f"      Righe dopo pulizia: {len(df_clean):,}")
print(f"      Periodo: {df_clean['InvoiceDate'].min().date()} → {df_clean['InvoiceDate'].max().date()}")

# ── 2. INSIGHT PRINCIPALI ───────────────────────────────────────────────────
print("\n[3/5] Calcolo insight...")

total_revenue = df_clean["Revenue"].sum()
total_orders = df_clean["InvoiceNo"].nunique()
total_customers = df_clean["CustomerID"].nunique()
avg_order_value = total_revenue / total_orders

print(f"\n  📊 RIEPILOGO GENERALE")
print(f"  {'Revenue totale:':30} £{total_revenue:>12,.2f}")
print(f"  {'Ordini unici:':30} {total_orders:>12,}")
print(f"  {'Clienti unici:':30} {total_customers:>12,}")
print(f"  {'Valore medio ordine:':30} £{avg_order_value:>12,.2f}")

# Top 10 prodotti per revenue
top_products = (
    df_clean.groupby("Description")["Revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
    .reset_index()
)
print(f"\n  🏆 TOP 5 PRODOTTI PER REVENUE:")
for _, row in top_products.head(5).iterrows():
    print(f"     • {row['Description'][:45]:46} £{row['Revenue']:>10,.2f}")

# Top paesi
top_countries = (
    df_clean.groupby("Country")["Revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(5)
    .reset_index()
)
print(f"\n  🌍 TOP 5 PAESI PER REVENUE:")
for _, row in top_countries.iterrows():
    print(f"     • {row['Country']:25} £{row['Revenue']:>12,.2f}")

# ── 3. GRAFICI ──────────────────────────────────────────────────────────────
print("\n[4/5] Generazione grafici...")

fig, axes = plt.subplots(2, 2, figsize=(16, 11))
fig.suptitle("Analisi Esplorativa – Vendite E-commerce (2010-2011)", fontsize=16, fontweight="bold", y=1.01)

# — Grafico 1: Revenue mensile
monthly = df_clean.groupby("Month")["Revenue"].sum().reset_index()
monthly["Month_str"] = monthly["Month"].astype(str)
ax1 = axes[0, 0]
ax1.bar(monthly["Month_str"], monthly["Revenue"], color="#1F4E79", edgecolor="white")
ax1.set_title("Revenue Mensile", fontweight="bold")
ax1.set_xlabel("Mese")
ax1.set_ylabel("Revenue (£)")
ax1.tick_params(axis="x", rotation=45)
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"£{x/1000:.0f}k"))

# — Grafico 2: Top 10 prodotti
ax2 = axes[0, 1]
colors = sns.color_palette("Blues_d", len(top_products))[::-1]
bars = ax2.barh(top_products["Description"].str[:35], top_products["Revenue"], color=colors)
ax2.set_title("Top 10 Prodotti per Revenue", fontweight="bold")
ax2.set_xlabel("Revenue (£)")
ax2.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"£{x/1000:.0f}k"))
ax2.invert_yaxis()

# — Grafico 3: Ordini per ora del giorno
hourly = df_clean.groupby("Hour")["InvoiceNo"].nunique().reset_index()
hourly.columns = ["Hour", "Orders"]
ax3 = axes[1, 0]
ax3.plot(hourly["Hour"], hourly["Orders"], marker="o", color="#1F4E79", linewidth=2.5, markersize=5)
ax3.fill_between(hourly["Hour"], hourly["Orders"], alpha=0.15, color="#1F4E79")
ax3.set_title("Ordini per Ora del Giorno", fontweight="bold")
ax3.set_xlabel("Ora")
ax3.set_ylabel("N° Ordini")
ax3.set_xticks(range(0, 24))

# — Grafico 4: Revenue per paese (escluso UK)
top_countries_no_uk = (
    df_clean[df_clean["Country"] != "United Kingdom"]
    .groupby("Country")["Revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(8)
    .reset_index()
)
ax4 = axes[1, 1]
colors2 = sns.color_palette("Blues_d", len(top_countries_no_uk))[::-1]
ax4.barh(top_countries_no_uk["Country"], top_countries_no_uk["Revenue"], color=colors2)
ax4.set_title("Top 8 Paesi per Revenue (escluso UK)", fontweight="bold")
ax4.set_xlabel("Revenue (£)")
ax4.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"£{x/1000:.0f}k"))
ax4.invert_yaxis()

plt.tight_layout()
plt.savefig("analisi_vendite.png", bbox_inches="tight")
plt.show()
print("      ✅ Grafico salvato: analisi_vendite.png")

# ── 4. EXPORT ────────────────────────────────────────────────────────────────
print("\n[5/5] Export dati...")
top_products.to_csv("top_prodotti.csv", index=False)
monthly.to_csv("revenue_mensile.csv", index=False)
print("      ✅ top_prodotti.csv")
print("      ✅ revenue_mensile.csv")

print("\n" + "=" * 60)
print("ANALISI COMPLETATA")
print("=" * 60)
