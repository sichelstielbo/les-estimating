import streamlit as st
import pandas as pd
import altair as alt

# --- Page config ---
st.set_page_config(
    page_title="LES Estimating Dashboard",
    page_icon="🏗️",
    layout="wide"
)

# --- Load data ---
@st.cache_data
def load_data():
    return pd.read_csv("/Users/bosichelstiel/Desktop/LES Docs/Master_Historical_Database.csv")

df = load_data()

# --- Sidebar filters ---
st.sidebar.title("Filters")
st.sidebar.markdown("---")

all_pms     = sorted(df['PM'].unique())
all_scopes  = sorted(df['Scope'].unique())
all_states  = sorted(df['State'].unique())

selected_pms    = st.sidebar.multiselect("Project Manager", all_pms,    default=all_pms)
selected_scopes = st.sidebar.multiselect("Scope",           all_scopes, default=all_scopes)
selected_states = st.sidebar.multiselect("State",           all_states, default=all_states)

st.sidebar.markdown("---")
st.sidebar.caption("Run les_extract.py to refresh data")

# --- Apply filters ---
filtered = df[
    df['PM'].isin(selected_pms) &
    df['Scope'].isin(selected_scopes) &
    df['State'].isin(selected_states)
]

analysis = filtered[filtered['Original_Est'] > 0]

# --- Header ---
st.title("LE Schwartz & Sons")
st.subheader("Estimating & Job Cost Dashboard")
st.markdown("---")

# --- KPI Metrics ---
total_jobs  = filtered['Job'].nunique()
total_est   = analysis['Original_Est'].sum()
total_proj  = analysis['Projected_Cost'].sum()
total_var   = total_est - total_proj
var_pct     = (total_var / total_est * 100) if total_est > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Jobs",            f"{total_jobs}")
col2.metric("Total Estimated", f"${total_est:,.0f}")
col3.metric("Total Projected", f"${total_proj:,.0f}")
col4.metric("Variance $",      f"${total_var:+,.0f}")
col5.metric("Variance %",      f"{var_pct:+.1f}%")

st.markdown("---")

# --- Jobs Table ---
st.subheader("Jobs Breakdown")

job_summary = (
    analysis.groupby(['Job', 'PM', 'City', 'State'])
    .agg(
        Estimated  =('Original_Est',   'sum'),
        Projected  =('Projected_Cost', 'sum'),
        Actual     =('Actual_to_Date', 'sum'),
    )
    .round(0)
    .reset_index()
)
job_summary['Variance $'] = job_summary['Estimated'] - job_summary['Projected']
job_summary['Variance %'] = (job_summary['Variance $'] / job_summary['Estimated'] * 100).round(1)
job_summary = job_summary.sort_values('Variance %')

def color_variance(val):
    if isinstance(val, str) and '%' in val:
        num = float(val.replace('%','').replace('+',''))
        color = '#1E8449' if num >= 0 else '#C0392B'
        return f'color: {color}; font-weight: bold'
    return ''

job_summary['Variance %'] = job_summary['Variance %'].apply(lambda x: f"{x:+.1f}%")
job_summary['Estimated']  = job_summary['Estimated'].apply(lambda x: f"${x:,.0f}")
job_summary['Projected']  = job_summary['Projected'].apply(lambda x: f"${x:,.0f}")
job_summary['Actual']     = job_summary['Actual'].apply(lambda x: f"${x:,.0f}")
job_summary['Variance $'] = job_summary['Variance $'].apply(lambda x: f"${x:+,.0f}")

styled = job_summary.style.map(color_variance, subset=['Variance %'])
st.dataframe(styled, use_container_width=True, hide_index=True)

st.markdown("---")

# --- Charts ---
left, right = st.columns(2)

with left:
    st.subheader("Variance by Scope")
    scope_chart = (
        analysis.groupby('Scope')
        .agg(Est=('Original_Est','sum'), Proj=('Projected_Cost','sum'))
        .reset_index()
    )
    scope_chart['Variance %'] = ((scope_chart['Est'] - scope_chart['Proj']) / scope_chart['Est'] * 100).round(1)
    scope_chart = scope_chart.sort_values('Variance %')
    scope_chart['Color'] = scope_chart['Variance %'].apply(lambda x: 'Over Budget' if x < 0 else 'Under Budget')

    chart = alt.Chart(scope_chart).mark_bar().encode(
        x=alt.X('Variance %:Q', title='Variance %'),
        y=alt.Y('Scope:N', sort='-x', title=''),
        color=alt.Color('Color:N', scale=alt.Scale(
            domain=['Over Budget', 'Under Budget'],
            range=['#C0392B', '#1E8449']
        ), legend=alt.Legend(title='')),
        tooltip=['Scope', 'Variance %']
    ).properties(height=400)
    st.altair_chart(chart, use_container_width=True)

with right:
    st.subheader("Variance by Category")
    cat_chart = (
        analysis.groupby('Category')
        .agg(Est=('Original_Est','sum'), Proj=('Projected_Cost','sum'))
        .reset_index()
    )
    cat_chart['Variance %'] = ((cat_chart['Est'] - cat_chart['Proj']) / cat_chart['Est'] * 100).round(1)
    cat_chart = cat_chart.sort_values('Variance %')
    cat_chart['Color'] = cat_chart['Variance %'].apply(lambda x: 'Over Budget' if x < 0 else 'Under Budget')

    chart2 = alt.Chart(cat_chart).mark_bar().encode(
        x=alt.X('Variance %:Q', title='Variance %'),
        y=alt.Y('Category:N', sort='-x', title=''),
        color=alt.Color('Color:N', scale=alt.Scale(
            domain=['Over Budget', 'Under Budget'],
            range=['#C0392B', '#1E8449']
        ), legend=alt.Legend(title='')),
        tooltip=['Category', 'Variance %']
    ).properties(height=400)
    st.altair_chart(chart2, use_container_width=True)

st.markdown("---")

# --- PM Summary ---
st.subheader("Summary by Project Manager")
pm_summary = (
    analysis.groupby('PM')
    .agg(
        Jobs      =('Job',           'nunique'),
        Estimated =('Original_Est',  'sum'),
        Projected =('Projected_Cost','sum'),
    )
    .reset_index()
    .round(0)
)
pm_summary['Variance $'] = pm_summary['Estimated'] - pm_summary['Projected']
pm_summary['Variance %'] = (pm_summary['Variance $'] / pm_summary['Estimated'] * 100).round(1)
pm_summary = pm_summary.sort_values('Variance %')

pm_summary['Variance %'] = pm_summary['Variance %'].apply(lambda x: f"{x:+.1f}%")
pm_summary['Estimated']  = pm_summary['Estimated'].apply(lambda x: f"${x:,.0f}")
pm_summary['Projected']  = pm_summary['Projected'].apply(lambda x: f"${x:,.0f}")
pm_summary['Variance $'] = pm_summary['Variance $'].apply(lambda x: f"${x:+,.0f}")

styled_pm = pm_summary.style.map(color_variance, subset=['Variance %'])
st.dataframe(styled_pm, use_container_width=True, hide_index=True)