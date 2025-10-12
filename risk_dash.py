# streamlit_grc_dashboard_final.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from datetime import datetime, date

# Import ReportLab for PDF Generation
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="GRC Dashboard",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- HELPER & STYLING FUNCTIONS ---

def apply_custom_styling():
    """Applies custom CSS for KPI cards and general styling."""
    st.markdown("""
    <style>
        .card {
            background-color: #2a2a3e; border-radius: 10px; padding: 15px;
            margin: 5px 0; box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2); color: white;
        }
        .kpi-title { font-size: 14px; font-weight: bold; color: #a9a9b3; }
        .kpi-value { font-size: 28px; font-weight: bold; color: #ffffff; }
    </style>
    """, unsafe_allow_html=True)

def render_kpi_card(title, value):
    """Renders a single KPI card."""
    st.markdown(f'<div class="card"><div class="kpi-title">{title}</div><div class="kpi-value">{value}</div></div>', unsafe_allow_html=True)

@st.cache_data
def read_data(uploaded_file):
    """Reads data from uploaded CSV or Excel file."""
    try:
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file)
        else:
            return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

def coerce_numeric(df, cols):
    """Converts specified columns to numeric, coercing errors."""
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    return df

def create_risk_matrix(df):
    """Creates an advanced risk matrix showing counts in each box."""
    if df.empty or 'Impact' not in df.columns or 'Likelihood' not in df.columns:
        fig = go.Figure()
        fig.update_layout(title='<b>Risk Matrix (Risk Counts)</b>', template='plotly_dark',
                          xaxis=dict(title='Impact', visible=False), yaxis=dict(title='Likelihood', visible=False))
        fig.add_annotation(text="No data to display", xref="paper", yref="paper", showarrow=False, font=dict(size=20))
        return fig

    agg_df = df.groupby(['Impact', 'Likelihood']).size().reset_index(name='count')
    text_matrix = [[0] * 5 for _ in range(5)]
    for index, row in agg_df.iterrows():
        try:
            impact_idx = int(row['Impact']) - 1
            likelihood_idx = int(row['Likelihood']) - 1
            if 0 <= impact_idx < 5 and 0 <= likelihood_idx < 5:
                text_matrix[likelihood_idx][impact_idx] = row['count']
        except (ValueError, TypeError):
            continue

    heatmap_z = [[1, 1, 2, 3, 3], [1, 2, 2, 3, 4], [2, 2, 3, 4, 4], [2, 3, 3, 4, 5], [3, 3, 4, 5, 5]]
    colorscale = [[0, 'rgb(12,128,64)'], [0.25, 'rgb(12,128,64)'], [0.25, 'rgb(255,255,0)'], [0.5, 'rgb(255,255,0)'],
                  [0.5, 'rgb(255,165,0)'], [0.75, 'rgb(255,165,0)'], [0.75, 'rgb(255,0,0)'], [1, 'rgb(255,0,0)']]
    heatmap_z_display = heatmap_z[::-1]
    text_matrix_display = text_matrix[::-1]
    fig = go.Figure(data=go.Heatmap(
        z=heatmap_z_display, x=[1, 2, 3, 4, 5], y=[1, 2, 3, 4, 5],
        colorscale=colorscale, showscale=False, text=text_matrix_display,
        texttemplate="%{text}", textfont={"size": 16, "color": "black"}
    ))
    fig.update_layout(
        title='<b>Risk Matrix (Risk Counts)</b>', template='plotly_dark', showlegend=False,
        xaxis=dict(tickmode='array', tickvals=[1, 2, 3, 4, 5], ticktext=['1-Low', '2-Minor', '3-Moderate', '4-Major', '5-Severe'], range=[0.5, 5.5], title='Impact'),
        yaxis=dict(tickmode='array', tickvals=[1, 2, 3, 4, 5], ticktext=['1-Rare', '2-Unlikely', '3-Possible', '4-Likely', '5-Almost Certain'], range=[0.5, 5.5], title='Likelihood'),
    )
    return fig

def generate_pdf_report(df):
    """Generates a PDF report from a DataFrame."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()
    title = Paragraph("GRC Risk Report", styles['h1'])
    report_date = Paragraph(f"Generated on: {date.today().strftime('%Y-%m-%d')}", styles['Normal'])
    elements.extend([title, report_date, Spacer(1, 24)])

    if not df.empty:
        kpi_text = f"""
        <b>Total Filtered Risks:</b> {len(df)} | 
        <b>Open Risks:</b> {int((df['Status'] == 'Open').sum())} | 
        <b>Avg. Likelihood:</b> {df['Likelihood'].mean():.2f} | 
        <b>Avg. Impact:</b> {df['Impact'].mean():.2f}
        """
        kpi_paragraph = Paragraph(kpi_text, styles['Normal'])
        elements.extend([kpi_paragraph, Spacer(1, 24)])
        report_df = df[['Risk ID', 'Title', 'Risk Owner', 'Risk Category', 'Likelihood', 'Impact', 'Risk Score', 'Status']].copy()
        report_df.rename(columns={'Risk Category': 'Category', 'Risk Score': 'Score'}, inplace=True)
        data = [report_df.columns.to_list()] + report_df.values.tolist()
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12), ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        table = Table(data, hAlign='LEFT')
        table.setStyle(table_style)
        elements.append(table)
    else:
        elements.append(Paragraph("No data available for the selected filters.", styles['Normal']))

    doc.build(elements)
    buffer.seek(0)
    return buffer

@st.cache_data
def to_excel(df):
    """Converts a DataFrame to an in-memory Excel file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Filtered Risks')
    return output.getvalue()

# --- MAIN APP LOGIC ---
def main():
    apply_custom_styling()
    st.sidebar.header("Upload Your Data")
    uploaded = st.sidebar.file_uploader("Upload risks CSV/Excel", type=["csv", "xlsx", "xls"])

    # Initialize session state for the dataframe if it doesn't exist
    if 'df' not in st.session_state:
        st.session_state.df = pd.DataFrame()

    if uploaded:
        st.session_state.df = read_data(uploaded)
        # Clear download states when a new file is uploaded
        st.session_state.excel_data = None
        st.session_state.pdf_data = None


    df = st.session_state.df

    with st.sidebar.expander("Register a New Risk"):
        with st.form("new_risk_form", clear_on_submit=True):
            st.write("Enter the details for the new risk:")
            title = st.text_input("Risk Title")
            owner = st.selectbox("Risk Owner (Form)", options=["IT", "Security", "Compliance", "Operations", "Finance"], key="form_owner")
            category = st.selectbox("Risk Category (Form)", options=["Data Protection", "Third-party", "Configuration", "Access Control", "Business Continuity"], key="form_category")
            likelihood = st.slider("Likelihood", 1, 5, 3, key="form_likelihood")
            impact = st.slider("Impact", 1, 5, 3, key="form_impact")
            status = st.selectbox("Status (Form)", options=["Open", "Mitigated", "Accepted", "Closed"], key="form_status")
            submitted = st.form_submit_button("Add Risk")
            
            if submitted and title:
                new_risk_id = f"R-{len(df) + 1001}"
                risk_score = likelihood * impact
                new_risk = pd.DataFrame([{"Risk ID": new_risk_id, "Title": title, "Risk Owner": owner,
                                          "Risk Category": category, "Likelihood": likelihood, "Impact": impact,
                                          "Risk Score": risk_score, "Status": status, "Control Effectiveness": "Medium",
                                          "Last Updated": date.today()}])
                st.session_state.df = pd.concat([st.session_state.df, new_risk], ignore_index=True)
                st.success(f"Risk '{title}' registered successfully!")

    st.sidebar.markdown("---")
    st.sidebar.header("Filters")
    
    if df.empty:
        st.info("👋 Welcome! Please upload a risk register file or register a new risk.")
        expected_columns = ["Risk ID", "Title", "Risk Owner", "Risk Category", "Likelihood", "Impact", 
                            "Risk Score", "Status", "Control Effectiveness", "Last Updated"]
        df = pd.DataFrame(columns=expected_columns)
        st.session_state.df = df
    
    if not df.empty:
        df = coerce_numeric(df, ['Likelihood', 'Impact'])
        if 'Risk Score' not in df.columns or df['Risk Score'].isna().all():
            if 'Likelihood' in df.columns and 'Impact' in df.columns:
                df['Risk Score'] = (df['Likelihood'].fillna(0) * df['Impact'].fillna(0))
            else:
                df['Risk Score'] = 0

    with st.sidebar:
        if not df.empty:
            owners = ['All'] + sorted(df['Risk Owner'].dropna().unique().tolist())
            categories = ['All'] + sorted(df['Risk Category'].dropna().unique().tolist())
            statuses = ['All'] + sorted(df['Status'].dropna().unique().tolist())
            owner_sel = st.selectbox("Risk Owner", owners)
            cat_sel = st.selectbox("Risk Category", categories)
            status_sel = st.selectbox("Status", statuses)
            
            df['Last Updated'] = pd.to_datetime(df['Last Updated'], errors='coerce').dt.date
            min_date, max_date = df['Last Updated'].min(), df['Last Updated'].max()
            date_range = st.date_input("Last Updated range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
            
            score_min, score_max = int(df['Risk Score'].min()), int(df['Risk Score'].max())
            score_sel = st.slider("Risk Score range", score_min, score_max, (score_min, score_max))
        else:
            owner_sel, cat_sel, status_sel, date_range, score_sel = "All", "All", "All", (date.today(), date.today()), (0, 25)
            st.selectbox("Risk Owner", ["All"], disabled=True)
            st.selectbox("Risk Category", ["All"], disabled=True)
            st.selectbox("Status", ["All"], disabled=True)
            st.date_input("Last Updated range", date_range, disabled=True)
            st.slider("Risk Score range", 0, 25, score_sel, disabled=True)

    filtered = df.copy()
    if not df.empty:
        if owner_sel != 'All': filtered = filtered[filtered['Risk Owner'] == owner_sel]
        if cat_sel != 'All': filtered = filtered[filtered['Risk Category'] == cat_sel]
        if status_sel != 'All': filtered = filtered[filtered['Status'] == status_sel]
        if len(date_range) == 2:
            filtered = filtered[(filtered['Last Updated'] >= date_range[0]) & (filtered['Last Updated'] <= date_range[1])]
        filtered = filtered[(filtered['Risk Score'] >= score_sel[0]) & (filtered['Risk Score'] <= score_sel[1])]

    st.title("📊 GRC Dashboard")
    
    st.markdown("### Key Metrics")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card("Total Filtered Risks", len(filtered))
    with col2:
        render_kpi_card("Open Risks", int((filtered['Status'] == 'Open').sum()) if not filtered.empty else 0)
    with col3:
        render_kpi_card("Avg. Likelihood", f"{filtered['Likelihood'].mean():.2f}" if not filtered.empty else "0.00")
    with col4:
        render_kpi_card("Avg. Impact", f"{filtered['Impact'].mean():.2f}" if not filtered.empty else "0.00")
        
    st.markdown("---")

    st.markdown("### Visual Analysis")
    st.plotly_chart(create_risk_matrix(filtered), use_container_width=True)

    vis_col1, vis_col2 = st.columns(2)
    with vis_col1:
        st.markdown("##### Top Risk Categories")
        if not filtered.empty:
            fig_treemap = px.treemap(filtered, path=[px.Constant("All"), 'Risk Category'], values='Risk Score')
            fig_treemap.update_layout(template="plotly_dark", margin=dict(t=30, l=10, r=10, b=10))
            st.plotly_chart(fig_treemap, use_container_width=True)
    with vis_col2:
        st.markdown("##### Status Distribution")
        if not filtered.empty:
            status_counts = filtered['Status'].value_counts()
            fig_pie = px.pie(status_counts, names=status_counts.index, values=status_counts.values, hole=0.4)
            fig_pie.update_layout(template="plotly_dark")
            st.plotly_chart(fig_pie, use_container_width=True)
        
    st.markdown("---")

    st.markdown("### Detailed Risk Register & Reports")
    if not filtered.empty:
        # --- DOWNLOAD LOGIC FULLY REVISED ---
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            if st.button("📥 Export to Excel", use_container_width=True):
                # Prepare data and store in session state
                st.session_state.excel_data = to_excel(filtered)
                st.session_state.pdf_data = None # Clear other download state
        with col_dl2:
            if st.button("📄 Download PDF Report", use_container_width=True):
                # Prepare data and store in session state
                st.session_state.pdf_data = generate_pdf_report(filtered)
                st.session_state.excel_data = None # Clear other download state

        # Trigger download if data is available in session state
        if 'excel_data' in st.session_state and st.session_state.excel_data:
            st.download_button(
                label="Click to Finalize Excel Download", data=st.session_state.excel_data,
                file_name="filtered_risk_register.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key='excel_download_trigger'
            )
            # Clear state after triggering download
            st.session_state.excel_data = None

        if 'pdf_data' in st.session_state and st.session_state.pdf_data:
            st.download_button(
                label="Click to Finalize PDF Download", data=st.session_state.pdf_data,
                file_name="grc_report.pdf",
                mime="application/pdf",
                key='pdf_download_trigger'
            )
            # Clear state after triggering download
            st.session_state.pdf_data = None
            
        st.dataframe(filtered.sort_values('Risk Score', ascending=False))

    st.markdown("---")

    st.markdown("### Non-Mitigated Risks")
    non_mitigated_risks = filtered[~filtered['Status'].isin(['Mitigated', 'Closed'])].copy()

    if not non_mitigated_risks.empty:
        st.write("The following risks require attention. Check 'Reviewed' and click the button below to update their status to 'Mitigated'.")
        non_mitigated_risks.insert(0, 'Reviewed', False)
        display_columns = ['Reviewed', 'Risk ID', 'Title', 'Risk Owner', 'Impact', 'Likelihood', 'Risk Score', 'Status']
        columns_to_show = [col for col in display_columns if col in non_mitigated_risks.columns]
        
        edited_df = st.data_editor(
            non_mitigated_risks[columns_to_show].sort_values('Risk Score', ascending=False),
            column_config={"Reviewed": st.column_config.CheckboxColumn("Reviewed?", default=False)},
            disabled=non_mitigated_risks.columns.drop('Reviewed'),
            hide_index=True,
            key="risk_editor"
        )
        
        if st.button("Update Statuses from Checklist"):
            reviewed_risk_ids = edited_df[edited_df['Reviewed'] == True]['Risk ID'].tolist()
            if reviewed_risk_ids:
                main_df = st.session_state.df
                main_df.loc[main_df['Risk ID'].isin(reviewed_risk_ids), 'Status'] = 'Mitigated'
                st.session_state.df = main_df
                st.success(f"Updated status for {len(reviewed_risk_ids)} risk(s) to 'Mitigated'.")
                st.rerun()
            else:
                st.warning("No risks were selected for update.")
    else:
        st.write("No non-mitigated risks to display for the selected filters.")

if __name__ == "__main__":
    main()
