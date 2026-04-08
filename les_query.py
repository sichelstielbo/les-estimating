import pandas as pd
import argparse
import sys

# --- File path ---
CSV_FILE = "/Users/bosichelstiel/Desktop/LES Docs/Master_Historical_Database.csv"

# --- Load data ---
df = pd.read_csv(CSV_FILE)

# --- Available filter values (for reference) ---
def show_options():
    print("\nAvailable filter values:")
    print(f"\n  PMs:    {', '.join(sorted(df['PM'].unique()))}")
    print(f"  Scopes: {', '.join(sorted(df['Scope'].unique()))}")
    print(f"  States: {', '.join(sorted(df['State'].unique()))}")
    print(f"  Cities: {', '.join(sorted(df['City'].unique()))}")
    print(f"  Jobs:   {', '.join(sorted(df['Job'].unique()))}")
    print()

# --- Argument parsing ---
parser = argparse.ArgumentParser(description='Query the LES estimating database.')
parser.add_argument('--pm',     type=str, help='Filter by project manager name')
parser.add_argument('--scope',  type=str, help='Filter by scope (e.g. "Siding")')
parser.add_argument('--state',  type=str, help='Filter by state (e.g. "GA")')
parser.add_argument('--city',   type=str, help='Filter by city')
parser.add_argument('--job',    type=str, help='Filter by job name')
parser.add_argument('--options',action='store_true', help='Show all available filter values')

args = parser.parse_args()

# Show options and exit if requested
if args.options:
    show_options()
    sys.exit()

# Require at least one filter
if not any([args.pm, args.scope, args.state, args.city, args.job]):
    print("\nPlease provide at least one filter.")
    print("Example: python3 les_query.py --pm \"Nick Riner\" --scope \"Siding\"")
    print("         python3 les_query.py --options   (to see all available values)")
    sys.exit()

# --- Apply filters ---
filtered = df.copy()

if args.pm:
    filtered = filtered[filtered['PM'].str.contains(args.pm, case=False, na=False)]
if args.scope:
    filtered = filtered[filtered['Scope'].str.contains(args.scope, case=False, na=False)]
if args.state:
    filtered = filtered[filtered['State'].str.contains(args.state, case=False, na=False)]
if args.city:
    filtered = filtered[filtered['City'].str.contains(args.city, case=False, na=False)]
if args.job:
    filtered = filtered[filtered['Job'].str.contains(args.job, case=False, na=False)]

if filtered.empty:
    print("\nNo results found for those filters. Try --options to see available values.")
    sys.exit()

# --- Build summary ---
active_filters = []
if args.pm:    active_filters.append(f"PM: {args.pm}")
if args.scope: active_filters.append(f"Scope: {args.scope}")
if args.state: active_filters.append(f"State: {args.state}")
if args.city:  active_filters.append(f"City: {args.city}")
if args.job:   active_filters.append(f"Job: {args.job}")

print(f"\n{'='*60}")
print(f"QUERY: {' | '.join(active_filters)}")
print(f"{'='*60}")
print(f"Matched {len(filtered)} rows across "
      f"{filtered['Job'].nunique()} job(s)\n")

# --- Jobs breakdown ---
job_summary = (
    filtered[filtered['Original_Est'] > 0]
    .groupby(['Job', 'PM', 'City', 'State'])
    .agg(
        Total_Est  =('Original_Est',    'sum'),
        Total_Proj =('Projected_Cost',  'sum'),
        Total_Act  =('Actual_to_Date',  'sum'),
    )
    .round(2)
)
job_summary['Variance_%'] = (
    (job_summary['Total_Est'] - job_summary['Total_Proj'])
    / job_summary['Total_Est'] * 100
).round(1)

print(f"{'JOB':<30} {'PM':<16} {'ESTIMATED':>14} {'PROJECTED':>14} {'VARIANCE':>10}")
print("-" * 88)
for (job, pm, city, state), row in job_summary.iterrows():
    direction = "▲" if row['Variance_%'] >= 0 else "▼"
    print(f"{job:<30} {pm:<16} "
          f"${row['Total_Est']:>13,.0f} "
          f"${row['Total_Proj']:>13,.0f} "
          f"  {direction}{abs(row['Variance_%']):>5.1f}%")

# --- Totals ---
total_est  = filtered[filtered['Original_Est'] > 0]['Original_Est'].sum()
total_proj = filtered[filtered['Original_Est'] > 0]['Projected_Cost'].sum()
total_var  = (total_est - total_proj) / total_est * 100 if total_est > 0 else 0

print("-" * 88)
print(f"{'TOTAL':<30} {'':<16} "
      f"${total_est:>13,.0f} "
      f"${total_proj:>13,.0f} "
      f"  {'▲' if total_var >= 0 else '▼'}{abs(total_var):>5.1f}%")

# --- Category breakdown ---
print(f"\n{'CATEGORY BREAKDOWN':}")
print("-" * 50)
cat_summary = (
    filtered[filtered['Original_Est'] > 0]
    .groupby('Category')
    .agg(
        Total_Est  =('Original_Est',   'sum'),
        Total_Proj =('Projected_Cost', 'sum'),
    )
    .round(2)
    .sort_values('Total_Est', ascending=False)
)
cat_summary['Variance_%'] = (
    (cat_summary['Total_Est'] - cat_summary['Total_Proj'])
    / cat_summary['Total_Est'] * 100
).round(1)

for cat, row in cat_summary.iterrows():
    direction = "▲" if row['Variance_%'] >= 0 else "▼"
    print(f"  {cat:<20} Est: ${row['Total_Est']:>10,.0f}   "
          f"Proj: ${row['Total_Proj']:>10,.0f}   "
          f"{direction}{abs(row['Variance_%']):.1f}%")

print()