"""
PG&E Migration Tool Comparator — Streamlit App
Upload CloudMigrate.store, Matilda, and Concierto Excel files for side-by-side comparison.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
import io

st.set_page_config(page_title="PG&E Migration Tool Comparator", layout="wide", page_icon="🔄")

# ---- Custom CSS ----
st.markdown("""
<style>
    .main-header {font-size: 2.2rem; font-weight: bold; color: #1F4E79; text-align: center; margin-bottom: 0.5rem;}
    .sub-header {font-size: 1.1rem; color: #666; text-align: center; margin-bottom: 2rem;}
    .metric-card {background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%); padding: 1.2rem;
                  border-radius: 10px; color: white; text-align: center; margin: 0.5rem 0;}
    .metric-value {font-size: 2rem; font-weight: bold;}
    .metric-label {font-size: 0.85rem; opacity: 0.9;}
    .winner-badge {background: #E2EFDA; color: #006100; padding: 0.3rem 0.8rem;
                   border-radius: 20px; font-weight: bold; display: inline-block;}
    .stTabs [data-baseweb="tab-list"] {gap: 8px;}
    .stTabs [data-baseweb="tab"] {padding: 10px 20px; font-weight: 600;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">PG&E Cloud Migration — Tool Comparator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">415 Applications | 4,273 Servers | Upload Excel files from CloudMigrate.store, Matilda, and Concierto for side-by-side analysis</div>', unsafe_allow_html=True)


# ---- Helper Functions ----
def load_ai_factors(wb):
    """Extract AI acceleration factors from AI Assumptions tab."""
    ws = wb['AI Assumptions']
    factors = {}
    for r in range(6, 40):
        cat = ws.cell(row=r, column=1).value
        factor = ws.cell(row=r, column=2).value
        if cat and isinstance(factor, (int, float)) and 0 < factor < 1:
            factors[cat] = round((1 - factor) * 100, 1)
    return factors


def load_activities(wb):
    """Extract migration activities with days."""
    ws = wb['Migration Activities']
    activities = []
    current_phase = ""
    for r in range(6, ws.max_row + 1):
        col1 = ws.cell(row=r, column=1).value
        col2 = ws.cell(row=r, column=2).value
        if col1 and 'Phase' in str(col1):
            current_phase = str(col1).replace('[COMPLETED]', '').strip()
            continue
        if not col2 or col2 == 'TOTAL':
            continue
        aws_t = ws.cell(row=r, column=5).value
        aws_t = aws_t if isinstance(aws_t, (int, float)) else 0
        az_t = ws.cell(row=r, column=8).value
        az_t = az_t if isinstance(az_t, (int, float)) else 0
        azl_t = ws.cell(row=r, column=11).value
        azl_t = azl_t if isinstance(azl_t, (int, float)) else 0
        team = ws.cell(row=r, column=3).value or ""
        cm_activity = ws.cell(row=r, column=14).value or ""
        activities.append({
            'Phase': current_phase,
            'Activity': str(col2).replace('[DONE]', '').strip(),
            'Team': team,
            'AWS_Trad': aws_t,
            'Azure_Trad': az_t,
            'AzLocal_Trad': azl_t,
            'Tool_Activity': cm_activity,
        })
    return pd.DataFrame(activities)


def load_roles(wb):
    """Extract team roles from Cost Estimation."""
    ws = wb['Cost Estimation']
    roles = []
    for r in range(6, 35):
        team = ws.cell(row=r, column=1).value
        role = ws.cell(row=r, column=2).value
        hr = ws.cell(row=r, column=3).value
        fte = ws.cell(row=r, column=6).value
        if team and role and isinstance(hr, (int, float)) and 'TOTAL' not in str(team):
            roles.append({
                'Team': team, 'Role': role,
                'Hourly_Rate': hr, 'FTE': fte if isinstance(fte, (int, float)) else 0
            })
    return pd.DataFrame(roles)


def load_phases(wb):
    """Extract phase costs from Cost Estimation Section B."""
    ws = wb['Cost Estimation']
    phases = []
    for r in range(30, 50):
        phase = ws.cell(row=r, column=1).value
        if not phase or 'Section' in str(phase) or 'TOTAL' in str(phase):
            continue
        td = ws.cell(row=r, column=3).value
        ad = ws.cell(row=r, column=4).value
        ftes = ws.cell(row=r, column=5).value
        rate = ws.cell(row=r, column=6).value
        if isinstance(td, (int, float)):
            phases.append({
                'Phase': str(phase).replace('[COMPLETED]', '').strip(),
                'Trad_Days': td, 'AI_Days': ad if isinstance(ad, (int, float)) else 0,
                'FTEs': ftes if isinstance(ftes, (int, float)) else 0,
                'Rate': rate if isinstance(rate, (int, float)) else 0,
            })
    return pd.DataFrame(phases)


# ---- File Upload or Auto-Load Sample Data ----
import os

SAMPLE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data")
SAMPLE_FILES = {
    'CloudMigrate.store': 'CloudMigrate_Store.xlsx',
    'Matilda Cloud': 'Matilda_Cloud.xlsx',
    'Concierto Migrate': 'Concierto_Migrate.xlsx',
    'AWS Transform+Kiro': 'AWS_Transform_Kiro.xlsx',
}

st.sidebar.header("Data Source")
use_sample = st.sidebar.checkbox("Use pre-loaded sample data (all 4 tools)", value=True)

tools = {}

if use_sample:
    st.sidebar.success("Loading all 4 tool comparisons from sample data")
    for name, filename in SAMPLE_FILES.items():
        filepath = os.path.join(SAMPLE_DIR, filename)
        if os.path.exists(filepath):
            tools[name] = openpyxl.load_workbook(filepath)
            st.sidebar.markdown(f"  {name}")
else:
    st.sidebar.markdown("Upload your own Excel files:")
    cm_file = st.sidebar.file_uploader("CloudMigrate.store", type=['xlsx'], key='cm')
    mt_file = st.sidebar.file_uploader("Matilda Cloud", type=['xlsx'], key='mt')
    cc_file = st.sidebar.file_uploader("Concierto Migrate", type=['xlsx'], key='cc')
    aws_file = st.sidebar.file_uploader("AWS Transform+Kiro", type=['xlsx'], key='aws')

    if cm_file:
        tools['CloudMigrate.store'] = openpyxl.load_workbook(io.BytesIO(cm_file.read()))
        cm_file.seek(0)
    if mt_file:
        tools['Matilda Cloud'] = openpyxl.load_workbook(io.BytesIO(mt_file.read()))
        mt_file.seek(0)
    if cc_file:
        tools['Concierto Migrate'] = openpyxl.load_workbook(io.BytesIO(cc_file.read()))
        cc_file.seek(0)
    if aws_file:
        tools['AWS Transform+Kiro'] = openpyxl.load_workbook(io.BytesIO(aws_file.read()))
        aws_file.seek(0)

if not tools:
    st.info("Please upload at least one Excel file or enable sample data to begin.")
    st.stop()

# ---- Parse all data ----
all_factors = {}
all_activities = {}
all_roles = {}
all_phases = {}

for name, wb in tools.items():
    all_factors[name] = load_ai_factors(wb)
    all_activities[name] = load_activities(wb)
    all_roles[name] = load_roles(wb)
    all_phases[name] = load_phases(wb)

tool_names = list(tools.keys())
tool_colors = {'CloudMigrate.store': '#2E75B6', 'Matilda Cloud': '#ED7D31', 'Concierto Migrate': '#70AD47', 'AWS Transform+Kiro': '#FF9900'}

# ---- TAB LAYOUT ----
tabs = st.tabs([
    "Executive Summary",
    "AI Factor Comparison",
    "Timeline & Cost",
    "Phase Breakdown",
    "Team & FTE View",
    "Activity Detail",
    "Agentic AI",
])

# ======== TAB 1: Executive Summary ========
with tabs[0]:
    st.header("Executive Summary")

    # Compute stats for all tools
    tool_stats = {}
    for name in tool_names:
        factors = all_factors[name]
        avg_savings = sum(factors.values()) / len(factors) if factors else 0
        phases_df = all_phases[name]
        trad_total = phases_df['Trad_Days'].sum()
        ai_total = phases_df['AI_Days'].sum()
        trad_cost = (phases_df['Trad_Days'] * phases_df['FTEs'] * phases_df['Rate']).sum()
        ai_cost = (phases_df['AI_Days'] * phases_df['FTEs'] * phases_df['Rate']).sum()
        tool_stats[name] = {
            'savings': avg_savings, 'ai_days': ai_total, 'trad_days': trad_total,
            'ai_cost': ai_cost, 'trad_cost': trad_cost, 'days_saved': trad_total - ai_total,
        }

    # ---- KPI cards (compact) ----
    cols = st.columns(len(tool_names))
    for i, name in enumerate(tool_names):
        s = tool_stats[name]
        color = tool_colors.get(name, '#666')
        with cols[i]:
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, {color}, {color}CC);
                        padding: 1.2rem; border-radius: 12px; color: white; text-align: center;">
                <div style="font-size: 1.1rem; font-weight: bold;">{name}</div>
                <div style="font-size: 2.2rem; font-weight: bold; margin: 0.3rem 0;">{s['savings']:.0f}%</div>
                <div style="opacity: 0.85; font-size: 0.85rem;">Savings | {int(s['ai_days'])}d | ${s['ai_cost']:,.0f}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("")

    # ---- "Who Wins Where" using native Streamlit components ----
    st.subheader("Who Wins Where")

    areas = [
        ('Discovery & Assessment', ['Infrastructure Discovery', 'App Dependency Mapping', 'Performance Baseline']),
        ('Database Migration', ['Database Analysis', 'Schema Conversion']),
        ('Planning & Wave', ['Wave Planning', 'TCO / Cost Modeling', 'License Analysis']),
        ('IaC / Terraform', ['Terraform/IaC Generation', 'Landing Zone Design']),
        ('Containerization', ['Container Artifact Gen', 'Containerize Execution']),
        ('Code Modernization', ['Framework Upgrades', 'Middleware Migration']),
        ('Rehost Execution', ['Rehost Execution', 'Replatform Execution']),
        ('Testing (FR + NFR)', ['FR Testing (60%)', 'NFR Testing (30%)']),
        ('Security & Compliance', ['Compliance & Security', 'Vuln Discovery Scan', 'Vuln Patch Execution']),
        ('Cost Optimization', ['Right-Sizing', 'Monitoring Setup']),
        ('Multi-Cloud Coverage', []),
    ]

    coverage_scores = {
        'CloudMigrate.store': 100, 'Matilda Cloud': 100,
        'Concierto Migrate': 100, 'AWS Transform+Kiro': 26,
    }

    # Build winner data for chart
    winner_rows = []
    for area_name, categories in areas:
        if area_name == 'Multi-Cloud Coverage':
            scores = {n: coverage_scores.get(n, 0) for n in tool_names}
        else:
            scores = {}
            for name in tool_names:
                vals = [all_factors[name].get(c, 0) for c in categories if c in all_factors[name]]
                scores[name] = sum(vals) / len(vals) if vals else 0

        winner = max(scores, key=scores.get)
        sorted_scores = sorted(scores.values(), reverse=True)
        is_tie = len(sorted_scores) > 1 and (sorted_scores[0] - sorted_scores[1]) < 2

        for name in tool_names:
            winner_rows.append({
                'Area': area_name, 'Tool': name, 'Score': round(scores.get(name, 0), 1),
                'Winner': 'TIE' if is_tie else winner,
            })

    winner_df = pd.DataFrame(winner_rows)

    # Grouped bar chart — this is clean and renders perfectly
    fig_wins = px.bar(
        winner_df, x='Area', y='Score', color='Tool', barmode='group',
        color_discrete_map=tool_colors,
        title='Savings % by Area — Who Wins Where',
        text='Score',
    )
    fig_wins.update_traces(texttemplate='%{text:.0f}%', textposition='outside', textfont_size=9)
    fig_wins.update_layout(
        height=500, xaxis_tickangle=-30,
        legend=dict(orientation='h', y=1.15, x=0.5, xanchor='center'),
        yaxis_title='Savings %', xaxis_title='',
    )
    st.plotly_chart(fig_wins, use_container_width=True)

    # Winner summary table using st.dataframe with emojis
    st.subheader("Winner by Area")
    summary_rows = []
    for area_name, categories in areas:
        if area_name == 'Multi-Cloud Coverage':
            scores = {n: coverage_scores.get(n, 0) for n in tool_names}
        else:
            scores = {}
            for name in tool_names:
                vals = [all_factors[name].get(c, 0) for c in categories if c in all_factors[name]]
                scores[name] = sum(vals) / len(vals) if vals else 0

        winner = max(scores, key=scores.get)
        sorted_s = sorted(scores.values(), reverse=True)
        is_tie = len(sorted_s) > 1 and (sorted_s[0] - sorted_s[1]) < 2
        winner_label = "TIE" if is_tie else winner

        row = {'Area': area_name, 'Winner': winner_label}
        for name in tool_names:
            s = scores.get(name, 0)
            is_w = (name == winner and not is_tie)
            row[name] = f"{s:.0f}%"
        summary_rows.append(row)

    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # ---- Overall Winner ----
    multi_cloud_coverage = {
        'CloudMigrate.store': 1.0, 'Matilda Cloud': 1.0,
        'Concierto Migrate': 1.0, 'AWS Transform+Kiro': 0.26,
    }

    def effective_savings(name):
        raw = tool_stats[name]['savings']
        cov = multi_cloud_coverage.get(name, 1.0)
        return raw * cov + 59 * (1 - cov)

    best_tool = max(tool_names, key=effective_savings)
    best_eff = effective_savings(best_tool)

    st.success(f"**Best for PG&E Multi-Cloud: {best_tool}** — {best_eff:.0f}% effective savings across all 415 apps")
    st.warning("**Coverage Note:** PG&E = AWS (26%) + Azure (35%) + Azure Local (27%) + Other (12%). AWS-only tools cover 26% of workload. Effective savings weighted by workload split.")

# ======== TAB 2: AI Factor Comparison ========
with tabs[1]:
    st.header("AI Acceleration Factors — Side by Side")

    # Build comparison DataFrame
    all_cats = set()
    for f in all_factors.values():
        all_cats.update(f.keys())
    all_cats = sorted(all_cats)

    factor_df = pd.DataFrame({'Category': all_cats})
    for name in tool_names:
        factor_df[name] = factor_df['Category'].map(all_factors[name]).fillna(0)

    if len(tool_names) > 1:
        factor_df['Delta'] = factor_df[tool_names[0]] - factor_df[tool_names[1]]

    # Bar chart
    fig = go.Figure()
    for name in tool_names:
        fig.add_trace(go.Bar(
            name=name, x=factor_df['Category'], y=factor_df[name],
            marker_color=tool_colors.get(name, '#666'),
            text=factor_df[name].apply(lambda x: f'{x:.0f}%'),
            textposition='outside',
        ))
    fig.update_layout(
        title="AI Savings % by Category", barmode='group',
        yaxis_title="Savings %", height=600,
        xaxis_tickangle=-45, legend=dict(orientation='h', y=1.12),
    )
    st.plotly_chart(fig, use_container_width=True)

    # Detail table
    st.subheader("Detailed Factor Table")
    st.dataframe(factor_df, use_container_width=True, height=600)

# ======== TAB 3: Timeline & Cost ========
with tabs[2]:
    st.header("Timeline & Cost Comparison")

    col1, col2 = st.columns(2)

    with col1:
        # Timeline chart
        timeline_data = []
        for name in tool_names:
            phases_df = all_phases[name]
            trad = phases_df['Trad_Days'].sum()
            ai = phases_df['AI_Days'].sum()
            timeline_data.append({'Tool': 'Traditional', 'Days': trad, 'Source': name})
            timeline_data.append({'Tool': name, 'Days': ai, 'Source': name})

        tl_df = pd.DataFrame(timeline_data)
        # Show unique tools
        summary = []
        if tool_names:
            trad_days = all_phases[tool_names[0]]['Trad_Days'].sum()
            summary.append({'Tool': 'Traditional', 'Days': trad_days, 'Months': round(trad_days / 22, 1)})
        for name in tool_names:
            ai_days = all_phases[name]['AI_Days'].sum()
            summary.append({'Tool': name, 'Days': ai_days, 'Months': round(ai_days / 22, 1)})

        sum_df = pd.DataFrame(summary)
        colors = ['#C00000'] + [tool_colors.get(n, '#666') for n in tool_names]
        fig_t = px.bar(sum_df, x='Tool', y='Days', color='Tool',
                       color_discrete_sequence=colors,
                       text='Days', title='Remaining Duration (Days)')
        fig_t.update_traces(textposition='outside')
        fig_t.update_layout(showlegend=False, height=450)
        st.plotly_chart(fig_t, use_container_width=True)

    with col2:
        # Cost chart
        cost_summary = []
        if tool_names:
            p = all_phases[tool_names[0]]
            trad_cost = (p['Trad_Days'] * p['FTEs'] * p['Rate']).sum()
            cost_summary.append({'Tool': 'Traditional', 'Cost': trad_cost})
        for name in tool_names:
            p = all_phases[name]
            ai_cost = (p['AI_Days'] * p['FTEs'] * p['Rate']).sum()
            cost_summary.append({'Tool': name, 'Cost': ai_cost})

        cost_df = pd.DataFrame(cost_summary)
        fig_c = px.bar(cost_df, x='Tool', y='Cost', color='Tool',
                       color_discrete_sequence=colors,
                       text=cost_df['Cost'].apply(lambda x: f'${x:,.0f}'),
                       title='Phase Labor Cost ($)')
        fig_c.update_traces(textposition='outside')
        fig_c.update_layout(showlegend=False, height=450, yaxis_tickformat='$,.0f')
        st.plotly_chart(fig_c, use_container_width=True)

    # Savings comparison table
    st.subheader("Savings Summary")
    savings_rows = []
    for name in tool_names:
        p = all_phases[name]
        trad_d = p['Trad_Days'].sum()
        ai_d = p['AI_Days'].sum()
        trad_c = (p['Trad_Days'] * p['FTEs'] * p['Rate']).sum()
        ai_c = (p['AI_Days'] * p['FTEs'] * p['Rate']).sum()
        savings_rows.append({
            'Tool': name,
            'Traditional Days': int(trad_d),
            'AI Days': int(ai_d),
            'Days Saved': int(trad_d - ai_d),
            'Months Saved': round((trad_d - ai_d) / 22, 1),
            'Traditional Cost': f'${trad_c:,.0f}',
            'AI Cost': f'${ai_c:,.0f}',
            'Cost Saved': f'${trad_c - ai_c:,.0f}',
        })
    st.dataframe(pd.DataFrame(savings_rows), use_container_width=True)

# ======== TAB 4: Phase Breakdown ========
with tabs[3]:
    st.header("Phase-by-Phase Comparison")

    for name in tool_names:
        st.subheader(name)
        p = all_phases[name].copy()
        p['Trad_Cost'] = p['Trad_Days'] * p['FTEs'] * p['Rate']
        p['AI_Cost'] = p['AI_Days'] * p['FTEs'] * p['Rate']
        p['Savings'] = p['Trad_Cost'] - p['AI_Cost']
        p['Savings_%'] = (p['Savings'] / p['Trad_Cost'] * 100).fillna(0).round(1)
        st.dataframe(p[['Phase', 'Trad_Days', 'AI_Days', 'FTEs', 'Rate', 'Trad_Cost', 'AI_Cost', 'Savings', 'Savings_%']],
                     use_container_width=True)

    # Stacked comparison
    if len(tool_names) > 1:
        st.subheader("AI Days by Phase — All Tools")
        phase_cmp = []
        for name in tool_names:
            for _, row in all_phases[name].iterrows():
                phase_cmp.append({'Tool': name, 'Phase': row['Phase'], 'AI_Days': row['AI_Days']})
        phase_cmp_df = pd.DataFrame(phase_cmp)
        fig_p = px.bar(phase_cmp_df, x='Phase', y='AI_Days', color='Tool', barmode='group',
                       color_discrete_map=tool_colors,
                       title='AI-Accelerated Days by Phase')
        fig_p.update_layout(xaxis_tickangle=-30, height=500)
        st.plotly_chart(fig_p, use_container_width=True)

# ======== TAB 5: Team & FTE ========
with tabs[4]:
    st.header("Team & FTE Distribution")

    for name in tool_names:
        st.subheader(name)
        roles_df = all_roles[name]
        team_summary = roles_df.groupby('Team').agg(
            FTEs=('FTE', 'sum'),
            Avg_Rate=('Hourly_Rate', 'mean'),
            Monthly_Cost=('FTE', lambda x: sum(x * roles_df.loc[x.index, 'Hourly_Rate'] * 8 * 22)),
        ).reset_index()
        team_summary['Annual_Cost'] = team_summary['Monthly_Cost'] * 12
        st.dataframe(team_summary, use_container_width=True)

    # FTE pie
    if tool_names:
        name = tool_names[0]
        roles_df = all_roles[name]
        team_fte = roles_df.groupby('Team')['FTE'].sum().reset_index()
        fig_fte = px.pie(team_fte, names='Team', values='FTE',
                         title=f'FTE Distribution ({name})',
                         color_discrete_sequence=px.colors.qualitative.Set2)
        fig_fte.update_traces(textinfo='label+value+percent')
        st.plotly_chart(fig_fte, use_container_width=True)

# ======== TAB 6: Activity Detail ========
with tabs[5]:
    st.header("Activity-Level Detail")

    selected_tool = st.selectbox("Select Tool", tool_names)
    acts_df = all_activities[selected_tool]

    # Filter
    phases = acts_df['Phase'].unique().tolist()
    selected_phase = st.multiselect("Filter by Phase", phases, default=phases)
    filtered = acts_df[acts_df['Phase'].isin(selected_phase)]

    st.dataframe(filtered, use_container_width=True, height=600)

    # Show tool-specific activities
    st.subheader("Tool Accelerator Activities")
    tool_acts = filtered[filtered['Tool_Activity'].str.len() > 5][['Activity', 'Team', 'Tool_Activity']]
    st.dataframe(tool_acts, use_container_width=True)

# ======== TAB 7: Agentic AI ========
with tabs[6]:
    st.header("Agentic AI Capabilities")

    agentic_data = {
        'Capability': [
            'Autonomous AI Agents', 'ML Prediction Models', 'MCP Microservices',
            'Guardrails & Safety', 'LLM Integration', 'Agentic API Routes',
            'AI Chat/Streaming', 'Project Memory', 'Code Scanning Rules',
            'Zero-Downtime Cutover', 'VM Throughput/Month', 'Multi-Cloud Support',
        ],
        'CloudMigrate.store': [
            '10 (Claude-powered)', '4 (91% accuracy)', '8 dedicated', '5 controls',
            'Claude AI (Anthropic)', '39 routes', 'YES (SSE)', 'YES', 'Guardrails engine',
            'Orchestrated', 'Orchestrated', 'AWS + Azure + Azure Local',
        ],
        'Matilda Cloud': [
            '0', '0', '0', '0', 'None', '0', 'NO', 'NO', 'None',
            'Workflow-based', 'Standard', 'AWS + Azure + GCP',
        ],
        'Concierto Migrate': [
            '2 (partial)', '1 (Maximize)', '3 (partial)', '1 (code scan)',
            'Not disclosed', '~10', 'Zero-code UI', 'NO', '5000+ rules (Modernize)',
            'Zero-downtime (CloudMach)', '1000+ VMs/month', 'AWS + Azure + GCP',
        ],
        'AWS Transform+Kiro': [
            'Q Developer (agentic)', '0 (uses Q ML)', '0 (native services)', 'IAM/Config native',
            'Claude AI (via Kiro)', '~15 (Q+Kiro)', 'Kiro IDE chat', 'NO', 'Q Transform 5000+',
            'MGN zero-downtime', 'MGN high-throughput', 'AWS ONLY',
        ],
    }
    ai_df = pd.DataFrame(agentic_data)

    # Show only uploaded tools
    display_cols = ['Capability'] + [n for n in tool_names if n in ai_df.columns]
    st.dataframe(ai_df[display_cols], use_container_width=True, height=500)

    # Radar chart
    if len(tool_names) >= 2:
        categories = ['AI Agents', 'ML Models', 'MCP Services', 'Guardrails',
                       'LLM', 'API Routes', 'Chat/Stream', 'Memory']
        cm_scores = [10, 4, 8, 5, 5, 5, 5, 5]
        mt_scores = [0, 0, 0, 0, 0, 0, 0, 0]
        cc_scores = [2, 1, 3, 1, 1, 2, 3, 0]

        aws_scores = [4, 1, 0, 4, 4, 3, 4, 0]  # Q Developer agentic but AWS-only

        score_map = {
            'CloudMigrate.store': cm_scores,
            'Matilda Cloud': mt_scores,
            'Concierto Migrate': cc_scores,
            'AWS Transform+Kiro': aws_scores,
        }

        fig_radar = go.Figure()
        for name in tool_names:
            if name in score_map:
                fig_radar.add_trace(go.Scatterpolar(
                    r=score_map[name] + [score_map[name][0]],
                    theta=categories + [categories[0]],
                    fill='toself', name=name,
                    line_color=tool_colors.get(name, '#666'),
                ))
        fig_radar.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 10])),
            title="Agentic AI Capability Radar",
            height=500,
        )
        st.plotly_chart(fig_radar, use_container_width=True)

    st.subheader("Key Insight")
    st.success("""
    **CloudMigrate.store** is the only platform with true Agentic AI:
    10 autonomous Claude-powered agents, 4 ML models, 8 MCP microservices, and a Guardrails engine.
    Matilda has zero AI agents (rule-based only). Concierto has partial AI in discovery + modernization
    but no autonomous agents.
    """)

# ---- Footer ----
st.divider()
st.markdown("""
<div style="text-align: center; color: #999; font-size: 0.85rem;">
    PG&E Migration Tool Comparator | 415 Apps | 4,273 Servers | Powered by Infosys Cobalt
</div>
""", unsafe_allow_html=True)
