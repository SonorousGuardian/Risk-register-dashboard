# live_grc_dashboard.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import plotly.io as pio
from plotly.subplots import make_subplots
import io
import time
import json
import re
from datetime import datetime, date
from typing import Dict, Optional, Tuple

# Google Sheets integration
import gspread
from google.oauth2.service_account import Credentials

# ReportLab for PDF generation
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ReportLabImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

# Required for Excel image export
from openpyxl.drawing.image import Image as OpenpyxlImage


# --- CONFIGURATION ---
class Config:
    """Configuration constants for the dashboard"""
    PAGE_TITLE = "Advanced GRC Dashboard"
    PAGE_ICON = "🛡️"
    LAYOUT = "wide"
    
    RISK_OWNERS = ["IT", "Security", "Compliance", "Operations", "Finance", "HR", "Legal"]
    RISK_CATEGORIES = ["Data Protection", "Third-party", "Configuration", "Access Control", "Business Continuity", "Cybersecurity"]
    STATUS_OPTIONS = ["Open", "In Progress", "Mitigated", "Accepted", "Closed"]
    CONTROL_EFFECTIVENESS_OPTIONS = ["Low", "Medium", "High"]
    
    COLORS = {'low_risk': '#2eb82e', 'medium_risk': '#ffa500', 'high_risk': '#ff4d4f', 'critical_risk': '#b22222'}

st.set_page_config(page_title=Config.PAGE_TITLE, page_icon=Config.PAGE_ICON, layout=Config.LAYOUT)

# --- DATA MANAGER ---
class DataManager:
    @staticmethod
    @st.cache_resource(ttl=3600)
    def _get_gspread_client(creds_json: dict):
        try:
            scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(creds_json, scopes=scopes)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"Failed to authenticate with Google Sheets: {e}")
            return None

    @st.cache_data(ttl=30)
    def read_live_data(_self, creds_json: dict, sheet_url: str) -> pd.DataFrame:
        gc = _self._get_gspread_client(creds_json)
        if not gc: return pd.DataFrame()
        try:
            spreadsheet = gc.open_by_url(sheet_url)
            worksheet = spreadsheet.sheet1
            records = worksheet.get_all_records()
            return pd.DataFrame(records) if records else pd.DataFrame()
        except gspread.exceptions.SpreadsheetNotFound:
            st.error("Spreadsheet not found. Check the URL and sharing settings.")
            return pd.DataFrame()
        except Exception as e:
            st.error(f"Error reading live data: {e}")
            return pd.DataFrame()

    def update_live_data(self, creds_json: dict, sheet_url: str, risk_ids: list, new_status: str, df: pd.DataFrame = None):
        """Update multiple risks in Google Sheets efficiently."""
        gc = self._get_gspread_client(creds_json)
        if not gc: return

        try:
            spreadsheet = gc.open_by_url(sheet_url)
            worksheet = spreadsheet.sheet1
            headers = worksheet.row_values(1)
            risk_id_col = headers.index('Risk ID') + 1
            status_col = headers.index('Status') + 1
            last_updated_col = headers.index('Last Updated') + 1

            # Build exact Risk ID -> row mapping to avoid partial matches
            risk_id_values = worksheet.col_values(risk_id_col)
            id_to_row = {val: idx + 1 for idx, val in enumerate(risk_id_values) if val}

            # Prepare batch updates for exact matched IDs only
            updates = []
            today_str = date.today().strftime('%Y-%m-%d')
            for risk_id in risk_ids:
                row_num = id_to_row.get(risk_id)
                # Skip header row and non-existent IDs
                if not row_num or row_num == 1:
                    continue
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(row_num, status_col),
                    "values": [[new_status]]
                })
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(row_num, last_updated_col),
                    "values": [[today_str]]
                })

            if updates:
                worksheet.batch_update(updates)  # One API call
                st.cache_data.clear()
        except Exception as e:
            st.error(f"Failed to update sheet: {e}")




    def upload_df_to_gsheet(_self, creds_json: dict, sheet_url: str, df: pd.DataFrame):
        gc = _self._get_gspread_client(creds_json)
        if not gc: return False
        try:
            spreadsheet = gc.open_by_url(sheet_url)
            worksheet = spreadsheet.sheet1
            headers = worksheet.row_values(1)
            # Reorder columns to match sheet headers to avoid misalignment
            upload_cols = [h for h in headers if h in df.columns]
            df_to_upload = df[upload_cols].copy()
            rows_to_append = df_to_upload.astype(str).values.tolist()
            worksheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
            st.cache_data.clear()
            return True
        except Exception as e:
            st.error(f"Failed to upload data to Google Sheet: {e}"); return False

    @staticmethod
    def read_from_file(uploaded_file) -> pd.DataFrame:
        try:
            return pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Error reading file: {e}"); return pd.DataFrame()

    def preprocess_data(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        df_clean = df.copy()
        # Ensure Likelihood and Impact exist and are numeric
        for col in ['Likelihood', 'Impact']:
            if col not in df_clean.columns:
                df_clean[col] = 1
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(1).astype(int)
        df_clean['Risk Score'] = df_clean['Likelihood'] * df_clean['Impact']
        df_clean['Last Updated'] = pd.to_datetime(df_clean.get('Last Updated'), errors='coerce').dt.date
        for col, default in [('Status', 'Open'), ('Risk Owner', 'Unknown'), ('Control Effectiveness', 'Medium')]:
             if col not in df_clean.columns: df_clean[col] = default
             else: df_clean[col] = df_clean[col].fillna(default)
        return df_clean

    def filter_data(self, df: pd.DataFrame, filters: Dict) -> pd.DataFrame:
        if df.empty: return df
        query_parts, params = [], {}
        if filters.get('owner') != 'All': query_parts.append("`Risk Owner` == @owner"); params['owner'] = filters['owner']
        if filters.get('status') != 'All': query_parts.append("`Status` == @status"); params['status'] = filters['status']
        if filters.get('control') != 'All': query_parts.append("`Control Effectiveness` == @control"); params['control'] = filters['control']
        score_range = filters.get('score_range', (0, 25)); query_parts.append("`Risk Score` >= @score_min and `Risk Score` <= @score_max"); params.update({'score_min': score_range[0], 'score_max': score_range[1]})
        if not query_parts: return df
        return df.query(" and ".join(query_parts), local_dict=params, engine='python')

# --- VISUALIZATION & UI MANAGERS ---
class VisualizationManager:
    @staticmethod
    def create_risk_matrix(df: pd.DataFrame) -> go.Figure:
        if df.empty: return go.Figure().update_layout(title='Risk Matrix (No data)', template='plotly_dark')
        count_matrix = pd.pivot_table(df, index='Impact', columns='Likelihood', aggfunc='size', fill_value=0).reindex(index=range(1, 6), columns=range(1, 6), fill_value=0)
        risk_level_matrix = np.fromfunction(lambda i, j: (i + 1) * (j + 1), (5, 5), dtype=int)
        colorscale = [[0.0, '#2E7D32'], [0.25, '#FFEB3B'], [0.75, '#FF9800'], [1.0, '#B71C1C']]
        fig = go.Figure(data=go.Heatmap(z=risk_level_matrix, x=count_matrix.columns, y=count_matrix.index, colorscale=colorscale, hovertemplate="<b>Likelihood:</b> %{x}<br><b>Impact:</b> %{y}<br><b>Risk Level:</b> %{z}<br><b>Risks:</b> %{customdata}<extra></extra>", customdata=count_matrix.values, colorbar=dict(title="Level")))
        annotations = [dict(x=j, y=i, text=f"<b>{count_matrix.loc[i, j]}</b>", showarrow=False, font=dict(color="white" if (i * j) >= 10 else "black", size=18)) for i in count_matrix.index for j in count_matrix.columns]
        fig.update_layout(template='plotly_dark', title=dict(text="<b>Risk Heatmap</b> (Count of Risks)", x=0.5, font_size=20), xaxis=dict(title='<b>Likelihood</b>', side="bottom"), yaxis=dict(title='<b>Impact</b>'), height=600, annotations=annotations, margin=dict(l=40, r=40, t=80, b=40))
        return fig

    @staticmethod
    def create_distribution_charts(df: pd.DataFrame) -> go.Figure:
        if df.empty: return go.Figure().update_layout(title='No data for distribution', template='plotly_dark')
        fig = make_subplots(
            rows=1, cols=3,
            subplot_titles=("By Risk Owner", "By Category", "By Status")
        )
        for i, col in enumerate(['Risk Owner', 'Risk Category', 'Status']):
            counts = df[col].value_counts()
            fig.add_trace(go.Bar(x=counts.index, y=counts.values, name=col), row=1, col=i+1)
        
        fig.update_layout(title_text="Risk Distribution", showlegend=False, template='plotly_dark', height=400)
        return fig

    @staticmethod
    def create_control_effectiveness_chart(df: pd.DataFrame) -> go.Figure:
        if df.empty or 'Control Effectiveness' not in df.columns:
            return go.Figure().update_layout(title='No data for control analysis', template='plotly_dark')
        
        avg_scores = df.groupby('Control Effectiveness')['Risk Score'].mean().reset_index().sort_values('Risk Score', ascending=False)

        fig = px.bar(avg_scores, x='Control Effectiveness', y='Risk Score', 
                     title='Average Risk Score by Control Effectiveness',
                     color='Control Effectiveness',
                     color_discrete_map={
                         'Low': Config.COLORS['high_risk'],
                         'Medium': Config.COLORS['medium_risk'],
                         'High': Config.COLORS['low_risk']
                     },
                     category_orders={"Control Effectiveness": ["Low", "Medium", "High"]})
        fig.update_layout(template='plotly_dark', height=400, yaxis_title="Average Risk Score")
        return fig

class UIComponents:
    @staticmethod
    def apply_custom_styling():
        st.markdown("""<style>
            body { background-color: #0E1117; } .stApp { background-color: #0E1117; }
            .card { background: #161B22; border-radius: 10px; padding: 20px; margin: 10px 0; color: white; border-left: 5px solid; box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2); }
            .kpi-title { font-size: 16px; color: #e0e0e0; margin-bottom: 8px; font-weight: 500; }
            .kpi-value { font-size: 32px; font-weight: 700; margin: 0; }
            .risk-high { border-left-color: #ff4d4f; } .risk-medium { border-left-color: #ffa500; } .risk-low { border-left-color: #2eb82e; } .risk-critical { border-left-color: #b22222; }
            .stTabs [data-baseweb="tab-list"] { gap: 24px; } .stTabs [data-baseweb="tab"] { height: 50px; background-color: transparent; }
            .stTabs [aria-selected="true"] { background-color: #161B22; border-radius: 8px 8px 0 0; }
        </style>""", unsafe_allow_html=True)
    
    @staticmethod
    def render_kpi_card(title: str, value, risk_level: str = "low", icon: str = ""):
        st.markdown(f'<div class="card risk-{risk_level}"><div class="kpi-title">{icon} {title}</div><div class="kpi-value">{value}</div></div>', unsafe_allow_html=True)
    
    @staticmethod
    def get_risk_level(score: float) -> Tuple[str, str]:
        if score >= 20: return "critical", Config.COLORS['critical_risk']
        elif score >= 15: return "high", Config.COLORS['high_risk']
        elif score >= 8: return "medium", Config.COLORS['medium_risk']
        else: return "low", Config.COLORS['low_risk']

# --- IMPROVED REPORT MANAGER ---
class ReportManager:
    class GRCReportTemplate(SimpleDocTemplate):
        def __init__(self, filename, **kw):
            super().__init__(filename, **kw)
            self.pagesize = landscape(letter)

        def afterPage(self):
            canvas = self.canv
            canvas.saveState()
            canvas.setFont('Helvetica', 9)
            canvas.drawString(inch, 0.75 * inch, f"Page {self.page} | GRC Risk Report")
            canvas.drawRightString(self.width + self.leftMargin - inch, 0.75 * inch, f"Generated: {datetime.now():%Y-%m-%d}")
            canvas.restoreState()

    @staticmethod
    def generate_excel_report(df: pd.DataFrame, risk_matrix_fig: go.Figure, filters: Dict, session_mitigated_df: pd.DataFrame) -> bytes:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            report_info_df = pd.DataFrame({"Filter": list(filters.keys()), "Value": [str(v) for v in filters.values()]})
            report_info_df.to_excel(writer, index=False, sheet_name='Report_Info')
            df.to_excel(writer, index=False, sheet_name='Filtered_Risks')
            
            mitigated_df = df[df['Status'] == 'Mitigated'].sort_values('Risk Score', ascending=False)
            if not mitigated_df.empty:
                mitigated_df.to_excel(writer, index=False, sheet_name='All_Mitigated_Risks')
            
            if not session_mitigated_df.empty:
                session_mitigated_df.to_excel(writer, index=False, sheet_name='Mitigated_This_Session')

            img_bytes = pio.to_image(risk_matrix_fig, format="png", width=800, height=600, scale=2)
            matrix_ws = writer.book.create_sheet("Risk_Matrix")
            matrix_ws.add_image(OpenpyxlImage(io.BytesIO(img_bytes)), 'A1')

            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = max(len(str(cell.value)) for cell in col) + 2
        return output.getvalue()

    @staticmethod
    def generate_pdf_report(df: pd.DataFrame, risk_matrix_fig: go.Figure, filters: Dict, session_mitigated_df: pd.DataFrame) -> bytes:
        buffer = io.BytesIO()
        doc = ReportManager.GRCReportTemplate(buffer, pagesize=landscape(letter))
        styles = getSampleStyleSheet()
        elements = [Paragraph("GRC Risk Management Report", styles['Title']), Spacer(1, 12)]
        
        elements.append(Paragraph("<b>Active Filters</b>", styles['Heading2']))
        elements.append(Paragraph("<br/>".join([f"<b>{k.title()}:</b> {v}" for k,v in filters.items()]), styles['Normal']))
        elements.append(PageBreak())
        
        elements.append(Paragraph("<b>Risk Assessment Matrix</b>", styles['Heading2']))
        img_bytes = pio.to_image(risk_matrix_fig, format="png", width=550, height=412)
        elements.append(ReportLabImage(io.BytesIO(img_bytes)))
        elements.append(PageBreak())

        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#DCE6F1")),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor("#DCE6F1"), colors.white])
        ])
        
        report_cols = ['Risk ID', 'Title', 'Risk Owner', 'Risk Score', 'Status']

        non_mitigated_df = df[~df['Status'].isin(['Mitigated', 'Closed', 'Accepted'])].sort_values('Risk Score', ascending=False)
        if not non_mitigated_df.empty:
            elements.append(Paragraph("<b>Active (Non-Mitigated) Risks</b>", styles['Heading2']))
            data = [report_cols] + non_mitigated_df[report_cols].values.tolist()
            table = Table(data, colWidths=[inch, 4*inch, 1.5*inch, inch, inch])
            table.setStyle(table_style)
            elements.append(table)
            elements.append(PageBreak())

        if not session_mitigated_df.empty:
            elements.append(Paragraph("<b>Risks Mitigated This Session</b>", styles['Heading2']))
            data = [report_cols] + session_mitigated_df[report_cols].values.tolist()
            table = Table(data, colWidths=[inch, 4*inch, 1.5*inch, inch, inch])
            table.setStyle(table_style)
            elements.append(table)
            elements.append(Spacer(1, 24))

        mitigated_df = df[df['Status'] == 'Mitigated'].sort_values('Risk Score', ascending=False)
        if not mitigated_df.empty:
            elements.append(Paragraph("<b>All Mitigated Risks (in current view)</b>", styles['Heading2']))
            data = [report_cols] + mitigated_df[report_cols].values.tolist()
            table = Table(data, colWidths=[inch, 4*inch, 1.5*inch, inch, inch])
            table.setStyle(table_style)
            elements.append(table)

        doc.build(elements)
        return buffer.getvalue()


# --- MAIN APPLICATION ---
def main():
    st.title(f"{Config.PAGE_ICON} {Config.PAGE_TITLE}")
    UIComponents.apply_custom_styling()
    data_manager = DataManager()
    
    if 'last_filters' not in st.session_state: st.session_state.last_filters = {}
    if 'df' not in st.session_state: st.session_state.df = pd.DataFrame()
    if 'initial_mitigated_ids' not in st.session_state: st.session_state.initial_mitigated_ids = set()

    with st.sidebar:
        with st.expander("🔄 **Data Source**", expanded=True):
            data_source = st.radio("Choose source:", ("File Upload", "Google Sheets (Live)"), key="data_source", horizontal=True)
            if data_source == "Google Sheets (Live)":
                st.session_state.gsheet_url = st.text_input("Google Sheet URL", st.session_state.get('gsheet_url', ''))
                st.session_state.gsheet_creds_file = st.file_uploader("Upload Credentials JSON", type=['json'])
            else:
                uploaded_file = st.file_uploader("Upload Risk Register", type=['csv', 'xlsx', 'xls'])

    df_loaded = False
    creds_info = None
    if data_source == "Google Sheets (Live)":
        if st.session_state.gsheet_creds_file:
            st.session_state.gsheet_creds_file.seek(0)
            creds_info = json.load(st.session_state.gsheet_creds_file)
        
        if st.session_state.gsheet_url and creds_info:
            st.session_state.is_live = True
            raw_df = data_manager.read_live_data(creds_info, st.session_state.gsheet_url)
            if not raw_df.empty: 
                if 'df' not in st.session_state or not raw_df.equals(st.session_state.get('raw_df')):
                    st.session_state.raw_df = raw_df
                    st.session_state.df = data_manager.preprocess_data(raw_df)
                    st.session_state.initial_mitigated_ids = set(st.session_state.df[st.session_state.df['Status'] == 'Mitigated']['Risk ID'])
                df_loaded = True
    elif 'uploaded_file' in locals() and uploaded_file:
        st.session_state.is_live = False
        raw_df = data_manager.read_from_file(uploaded_file)
        if 'df' not in st.session_state or not raw_df.equals(st.session_state.get('raw_df')):
             st.session_state.raw_df = raw_df
             st.session_state.df = data_manager.preprocess_data(raw_df)
             st.session_state.initial_mitigated_ids = set(st.session_state.df[st.session_state.df['Status'] == 'Mitigated']['Risk ID'])
        df_loaded = True

    df = st.session_state.get('df', pd.DataFrame())

    with st.sidebar:
        with st.expander("➕ **Add New Risk**"):
            with st.form("new_risk_form", clear_on_submit=True):
                st.subheader("New Risk Details")
                title = st.text_input("Risk Title*")
                owner = st.selectbox("Risk Owner*", Config.RISK_OWNERS)
                category = st.selectbox("Risk Category*", Config.RISK_CATEGORIES)
                c1, c2 = st.columns(2)
                likelihood = c1.slider("Likelihood*", 1, 5, 3)
                impact = c2.slider("Impact*", 1, 5, 3)
                control = st.selectbox("Control Effectiveness*", Config.CONTROL_EFFECTIVENESS_OPTIONS, index=1)
                
                submitted = st.form_submit_button("Add Risk", type="primary", use_container_width=True)
                if submitted and title:
                    if not df_loaded:
                        st.warning("Please load data before adding a new risk.")
                    else:
                        new_risk_score = likelihood * impact
                        
                        if not df.empty and 'Risk ID' in df.columns:
                            numeric_ids = pd.to_numeric(df['Risk ID'].str.extract(r'R-(\d+)')[0], errors='coerce').dropna()
                            max_id = int(numeric_ids.max()) if not numeric_ids.empty else 999
                            new_id = f"R-{max_id + 1}"
                        else:
                            new_id = "R-1000"

                        new_row = pd.DataFrame([{'Risk ID': new_id, 'Title': title, 'Risk Owner': owner, 'Risk Category': category, 'Likelihood': likelihood, 'Impact': impact, 'Risk Score': new_risk_score, 'Status': 'Open', 'Control Effectiveness': control, 'Last Updated': date.today()}])

                        if st.session_state.is_live:
                            with st.spinner("Adding risk to Google Sheet..."):
                                data_manager.upload_df_to_gsheet(creds_info, st.session_state.gsheet_url, new_row)
                            st.success(f"Added risk {new_id} to Google Sheet!")
                        else:
                            st.session_state.df = pd.concat([df, data_manager.preprocess_data(new_row)], ignore_index=True)
                            st.success(f"Added risk {new_id} to session data!")
                        time.sleep(1)
                        st.rerun()

    if not df_loaded:
        st.info("👋 **Welcome!** Please configure a data source in the sidebar to begin."); return

    total_risks = len(df); mitigated_risks = len(df[df['Status'] == 'Mitigated'])
    avg_score = df['Risk Score'].mean(); critical_risks = len(df[df['Risk Score'] >= 20])
    level, _ = UIComponents.get_risk_level(avg_score)
    c1, c2, c3, c4 = st.columns(4)
    with c1: UIComponents.render_kpi_card("Total Risks", total_risks, level, "🗂️")
    with c2: UIComponents.render_kpi_card("Mitigated Risks", mitigated_risks, "low", "✅")
    with c3: UIComponents.render_kpi_card("Avg. Score", f"{avg_score:.1f}", level, "📈")
    with c4: UIComponents.render_kpi_card("Critical Risks (≥20)", critical_risks, "critical" if critical_risks > 0 else "low", "🚨")
    st.markdown("---")

    with st.sidebar:
        with st.expander("🔍 **Filters**", expanded=True):
            owners = ['All'] + sorted(df['Risk Owner'].dropna().unique())
            statuses = ['All'] + sorted(df['Status'].dropna().unique())
            controls = ['All'] + sorted(df['Control Effectiveness'].dropna().unique())
            owner_sel = st.selectbox("Owner", owners); status_sel = st.selectbox("Status", statuses)
            control_sel = st.selectbox("Control Effectiveness", controls)
            score_min, score_max = int(df['Risk Score'].min()), int(df['Risk Score'].max())
            score_sel = st.slider("Score Range", score_min, score_max, (score_min, score_max))
        
        current_filters = {'owner': owner_sel, 'status': status_sel, 'control': control_sel, 'score_range': score_sel}
        if st.session_state.last_filters != current_filters:
            if 'excel_report' in st.session_state: del st.session_state.excel_report
            if 'pdf_report' in st.session_state: del st.session_state.pdf_report
        st.session_state.last_filters = current_filters
        filtered_df = data_manager.filter_data(df, current_filters)

        if not st.session_state.is_live:
            with st.expander("📤 **Export Data**"):
                st.download_button("Download Updated CSV", data=df.to_csv(index=False).encode('utf-8'), file_name=f"updated_risks.csv", use_container_width=True)

    tab1, tab2, tab3, tab4 = st.tabs(["📋 **Register**", "📊 **Analytics**", "✅ **Checklist**", "📄 **Reports**"])

    with tab1: st.dataframe(filtered_df, use_container_width=True, height=500)

    with tab2:
        st.header("Risk Analytics Dashboard")
        st.plotly_chart(VisualizationManager.create_risk_matrix(filtered_df), use_container_width=True)
        st.plotly_chart(VisualizationManager.create_distribution_charts(filtered_df), use_container_width=True)
        st.plotly_chart(VisualizationManager.create_control_effectiveness_chart(filtered_df), use_container_width=True)

    
    with tab3:
        st.header("🔄 Risk Mitigation Checklist")

        # Initialize once
        if "checklist_df" not in st.session_state:
            df_current = st.session_state.df.copy()
            st.session_state.checklist_df = df_current[~df_current["Status"].isin(["Mitigated", "Closed"])].copy()
            st.session_state.checklist_df.insert(0, "Mitigate", False)

        checklist_df = st.session_state.checklist_df

        # --- Data Editor (outside form for live count updates) ---
        edited_df = st.data_editor(
            checklist_df[["Mitigate", "Risk ID", "Title", "Risk Score", "Status"]],
            use_container_width=True,
            hide_index=True,
            key="risk_checklist_editor",
            column_config={
                "Mitigate": st.column_config.CheckboxColumn("Select"),
                "Risk Score": st.column_config.ProgressColumn("Score", min_value=1, max_value=25, format="%d"),
            }
        )

        # Normalize types for reliable comparisons
        edited_df["Risk ID"] = edited_df["Risk ID"].astype(str)
        st.session_state.df["Risk ID"] = st.session_state.df["Risk ID"].astype(str)

        selected_ids = edited_df.loc[edited_df["Mitigate"], "Risk ID"].tolist()
        st.caption(f"Selected: {len(selected_ids)} risk(s)")

        # Submit only wraps the button, using current selection
        with st.form("risk_checklist_form", clear_on_submit=False):
            submit_mitigate = st.form_submit_button(
                "✅ Mitigate Selected",
                type="primary"
            )

        # Only update state if submitted
        if submit_mitigate:
            if not selected_ids:
                st.warning("Please select at least one risk to mitigate.")
            else:
                with st.spinner("Updating mitigated risks..."):
                    if st.session_state.is_live:
                        data_manager.update_live_data(
                            creds_info,
                            st.session_state.gsheet_url,
                            selected_ids,
                            "Mitigated",
                            st.session_state.df
                        )
                        # Fetch updated data from Google Sheet
                        st.session_state.df = data_manager.read_live_data(creds_info, st.session_state.gsheet_url)
                    else:
                        st.session_state.df.loc[
                            st.session_state.df["Risk ID"].isin(selected_ids), "Status"
                        ] = "Mitigated"

                # Remove selected risks from checklist immediately
                if "checklist_df" in st.session_state:
                    st.session_state.checklist_df = st.session_state.checklist_df[
                        ~st.session_state.checklist_df["Risk ID"].isin(selected_ids)
                    ].copy()
                
                # Rebuild checklist from authoritative df (ensures consistency)
                st.session_state.checklist_df = st.session_state.df[
                    ~st.session_state.df["Status"].isin(["Mitigated", "Closed"])
                ].copy()
                st.session_state.checklist_df.insert(0, "Mitigate", False)

                st.success(f"✅ {len(selected_ids)} risk(s) marked as mitigated.")
                st.rerun()

        # --- Progress and Metrics (update on submit/rerun) ---
        total_mitigatable = len(st.session_state.df[~st.session_state.df["Status"].isin(["Closed", "Accepted"])])
        already_mitigated_total = len(st.session_state.df[st.session_state.df["Status"] == "Mitigated"])
        newly_selected_count = len(selected_ids)
        progress_value = (already_mitigated_total) / total_mitigatable if total_mitigatable > 0 else 0
        st.progress(progress_value, text=f"Overall Mitigation Progress ({progress_value:.0%})")

        s1, s2, s3 = st.columns(3)
        s1.metric("Active Risks in View", len(st.session_state.checklist_df))
        s2.metric("Selected for Mitigation", newly_selected_count)
        s3.metric("High/Critical in View", len(st.session_state.checklist_df[st.session_state.checklist_df["Risk Score"] >= 15]))

        # Manual refresh button (optional)
        if st.button("🔄 Refresh Checklist"):
            with st.spinner("Refreshing data from Google Sheet..."):
                if st.session_state.is_live:
                    st.session_state.df = data_manager.read_live_data(creds_info, st.session_state.gsheet_url)
                # Remove mitigated/closed
                st.session_state.checklist_df = st.session_state.df[
                    ~st.session_state.df["Status"].isin(["Mitigated", "Closed"])
                ].copy()
                st.session_state.checklist_df.insert(0, "Mitigate", False)
                st.success("✅ Checklist refreshed successfully!")
                st.rerun()


            
        with tab4:
            st.header("📄 Reports & Exports")
            if not filtered_df.empty:
                current_mitigated_ids = set(df[df['Status'] == 'Mitigated']['Risk ID'])
                session_mitigated_ids = current_mitigated_ids - st.session_state.initial_mitigated_ids
                session_mitigated_df = df[df['Risk ID'].isin(session_mitigated_ids)]

                risk_matrix_fig = VisualizationManager.create_risk_matrix(filtered_df)
                c1, c2 = st.columns(2)

                with c1:
                    if st.button("📘 Generate Excel Report", use_container_width=True):
                        with st.spinner("Generating Excel report..."):
                            excel_bytes = ReportManager.generate_excel_report(filtered_df, risk_matrix_fig, current_filters, session_mitigated_df)
                            st.session_state.excel_report = excel_bytes
                        st.success("✅ Excel report generated successfully!")

                    if 'excel_report' in st.session_state:
                        st.download_button(
                            label="⬇️ Download Excel Report",
                            data=st.session_state.excel_report,
                            file_name=f"GRC_Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                with c2:
                    if st.button("📄 Generate PDF Report", use_container_width=True):
                        with st.spinner("Generating PDF report..."):
                            pdf_bytes = ReportManager.generate_pdf_report(filtered_df, risk_matrix_fig, current_filters, session_mitigated_df)
                            st.session_state.pdf_report = pdf_bytes
                        st.success("✅ PDF report generated successfully!")

                    if 'pdf_report' in st.session_state:
                        st.download_button(
                            label="⬇️ Download PDF Report",
                            data=st.session_state.pdf_report,
                            file_name=f"GRC_Report_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

            else:
                st.warning("⚠️ No data available for generating reports.")

if __name__ == "__main__":
    main()

