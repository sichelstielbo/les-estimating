import pandas as pd

FILES = [
    "/Users/bosichelstiel/Desktop/LES Docs/JCW- Jan 2026 Costing (NICK).xlsx",
    "/Users/bosichelstiel/Desktop/LES Docs/JCW- Jan 2026 Costing (MARK).xlsx",
    "/Users/bosichelstiel/Desktop/LES Docs/JCW- Jan 2025 Costing (SHANE).xlsx",
]

SKIP_TABS = ["Combined Data Set", "Pivot", "Sheet1"]

OUTPUT_FILE = "/Users/bosichelstiel/Desktop/LES Docs/Master_Historical_Database.csv"

def get_val(df, r, c):
    try:
        val = df.iloc[r, c]
        return val if pd.notna(val) else ""
    except:
        return ""

def extract_job_data(df):
    rows = []

    project = get_val(df, 0, 2)
    city    = get_val(df, 1, 2)
    state   = get_val(df, 1, 3)
    gc      = get_val(df, 2, 2)
    pm      = get_val(df, 3, 2)

    for idx, row in df.iterrows():
        if str(get_val(df, idx, 2)).strip().lower() == "original estimate":

            scope = str(get_val(df, idx - 1, 2)).strip()

            current = idx + 1
            while current < len(df):
                category = str(get_val(df, current, 1)).strip()

                if not category or category.lower() in ["nan", ""]:
                    break

                orig_est  = pd.to_numeric(get_val(df, current, 2), errors="coerce")
                proj_cost = pd.to_numeric(get_val(df, current, 3), errors="coerce")
                actual    = pd.to_numeric(get_val(df, current, 5), errors="coerce")

                orig_est  = 0 if pd.isna(orig_est)  else orig_est
                proj_cost = 0 if pd.isna(proj_cost) else proj_cost
                actual    = 0 if pd.isna(actual)    else actual

                if (orig_est != 0 or proj_cost != 0 or actual != 0) and scope.lower() != "total":
                    rows.append({
                        "Job":            project,
                        "City":           city,
                        "State":          state,
                        "GC":             gc,
                        "PM":             pm,
                        "Scope":          scope,
                        "Category":       category,
                        "Original_Est":   orig_est,
                        "Projected_Cost": proj_cost,
                        "Actual_to_Date": actual,
                        "Variance":       orig_est - proj_cost,
                    })

                current += 1

    return rows

# --- Run across all workbooks and all job tabs ---
all_rows = []

for filepath in FILES:
    filename = filepath.split("/")[-1]
    xls = pd.ExcelFile(filepath)

    job_tabs = [
        name for name in xls.sheet_names
        if not any(skip.lower() in name.lower() for skip in SKIP_TABS)
        and not name.lower().startswith("template")
    ]

    print(f"\n{filename}")
    for tab in job_tabs:
        df = pd.read_excel(xls, sheet_name=tab, header=None)
        rows = extract_job_data(df)
        all_rows.extend(rows)
        print(f"  {tab}: {len(rows)} rows")

# --- Save to CSV ---
master_df = pd.DataFrame(all_rows)

# --- Data quality fixes ---

# Fix PM name typos (extra spaces, formatting issues)
master_df['PM'] = master_df['PM'].str.strip().str.replace(r'\s+', ' ', regex=True)

# Fix specific PM name typos
PM_MAP = {
    'Mar k Coryell': 'Mark Coryell',
}
master_df['PM'] = master_df['PM'].replace(PM_MAP)

# Standardize scope names that mean the same thing
SCOPE_MAP = {
    'Gutters': 'Gutter',
}
master_df['Scope'] = master_df['Scope'].replace(SCOPE_MAP)

# Standardize state capitalization
master_df['State'] = master_df['State'].str.strip().str.upper()

# Fix city typos and trailing spaces
master_df['City'] = master_df['City'].str.strip()
master_df['Job']  = master_df['Job'].str.strip()
CITY_MAP = {
    'Altanta': 'Atlanta',
}
master_df['City'] = master_df['City'].replace(CITY_MAP)

master_df.to_csv(OUTPUT_FILE, index=False)

print(f"\nDone. {len(all_rows)} total rows saved to:")
print(OUTPUT_FILE)

# --- Summary Analysis ---
print("\n" + "="*60)
print("VARIANCE SUMMARY BY SCOPE")
print("="*60)

analysis_df = master_df[master_df["Original_Est"] > 0].copy()

scope_summary = (
    analysis_df.groupby("Scope")
    .agg(
        Jobs=("Job", "nunique"),
        Total_Original_Est=("Original_Est", "sum"),
        Total_Projected_Cost=("Projected_Cost", "sum"),
    )
    .round(2)
)

scope_summary["Variance_Pct"] = (
    (scope_summary["Total_Original_Est"] - scope_summary["Total_Projected_Cost"])
    / scope_summary["Total_Original_Est"] * 100
).round(1)

scope_summary = scope_summary.sort_values("Variance_Pct")

for scope, row in scope_summary.iterrows():
    direction = "OVER budget" if row["Variance_Pct"] < 0 else "under budget"
    print(f"\n{scope} ({int(row['Jobs'])} jobs)")
    print(f"  Total estimated:  ${row['Total_Original_Est']:>12,.0f}")
    print(f"  Total projected:  ${row['Total_Projected_Cost']:>12,.0f}")
    print(f"  Variance:         {row['Variance_Pct']:>+.1f}%  ({direction})")

print("\n" + "="*60)
print("VARIANCE SUMMARY BY PM")
print("="*60)

pm_summary = (
    analysis_df.groupby("PM")
    .agg(
        Jobs=("Job", "nunique"),
        Total_Original_Est=("Original_Est", "sum"),
        Total_Projected_Cost=("Projected_Cost", "sum"),
    )
    .round(2)
)

pm_summary["Variance_Pct"] = (
    (pm_summary["Total_Original_Est"] - pm_summary["Total_Projected_Cost"])
    / pm_summary["Total_Original_Est"] * 100
).round(1)

pm_summary = pm_summary.sort_values("Variance_Pct")

for pm, row in pm_summary.iterrows():
    direction = "OVER budget" if row["Variance_Pct"] < 0 else "under budget"
    print(f"\n{pm} ({int(row['Jobs'])} jobs)")
    print(f"  Total estimated:  ${row['Total_Original_Est']:>12,.0f}")
    print(f"  Total projected:  ${row['Total_Projected_Cost']:>12,.0f}")
    print(f"  Variance:         {row['Variance_Pct']:>+.1f}%  ({direction})")