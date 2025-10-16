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
from datetime import datetime, date
from typing import Dict, Optional, Tuple, List, Any

# Google Sheets integration
import gspread
from google.oauth2.service_account import Credentials
import re
from gspread.utils import rowcol_to_a1
from google.auth.exceptions import RefreshError

# ReportLab for PDF generation
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as ReportLabImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

# Required for Excel image export
try:
    from openpyxl.drawing.image import Image as OpenpyxlImage
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.warning("openpyxl not available - Excel export will be limited")

# --- CONFIGURATION ---
class Config:
    """Configuration constants for the dashboard"""
    PAGE_TITLE = "Risk Management Dashboard"
    PAGE_ICON = "🛡️"
    LAYOUT = "wide"
    
    RISK_OWNERS = ["IT", "Security", "Compliance", "Operations", "Finance", "HR", "Legal"]
    RISK_CATEGORIES = ["Data Protection", "Third-party", "Configuration", "Access Control", "Business Continuity", "Cybersecurity"]
    STATUS_OPTIONS = ["Open", "In Progress", "Mitigated", "Accepted", "Closed"]
    CONTROL_EFFECTIVENESS_OPTIONS = ["Low", "Medium", "High"]
    
    COLORS = {
        'low_risk': '#2eb82e', 
        'medium_risk': '#ffa500', 
        'high_risk': '#ff4d4f', 
        'critical_risk': '#b22222',
        'background': '#0E1117',
        'card_background': '#161B22'
    }
    
    # Google Sheets configuration
    REQUIRED_COLUMNS = ['Risk ID', 'Title', 'Risk Owner', 'Risk Category', 'Likelihood', 'Impact', 'Risk Score', 'Status', 'Control Effectiveness', 'Last Updated']

st.set_page_config(
    page_title=Config.PAGE_TITLE, 
    page_icon=Config.PAGE_ICON, 
    layout=Config.LAYOUT,
    initial_sidebar_state="expanded"
)

# --- IMPROVED DATA MANAGER ---
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

    @st.cache_data(ttl=30, show_spinner=False)
    def read_live_data(_self, creds_json: dict, sheet_url: str) -> pd.DataFrame:
        """Read data from Google Sheets with improved error handling"""
        if not creds_json or not sheet_url:
            return pd.DataFrame()
            
        gc = _self._get_gspread_client(creds_json)
        if not gc: 
            return pd.DataFrame()
            
        try:
            spreadsheet = gc.open_by_url(sheet_url)
            worksheet = spreadsheet.sheet1
            records = worksheet.get_all_records()
            
            if not records:
                st.warning("No data found in the Google Sheet")
                return pd.DataFrame()
                
            df = pd.DataFrame(records)
            
            # Validate required columns
            missing_cols = [col for col in Config.REQUIRED_COLUMNS if col not in df.columns]
            if missing_cols:
                st.warning(f"Missing columns in Google Sheet: {', '.join(missing_cols)}")
                
            return df
            
        except gspread.exceptions.SpreadsheetNotFound:
            st.error("Spreadsheet not found. Check the URL and sharing settings.")
            return pd.DataFrame()
        except gspread.exceptions.APIError as e:
            st.error(f"Google Sheets API error: {e}")
            return pd.DataFrame()
        except Exception as e:
            st.error(f"Unexpected error reading live data: {e}")
            return pd.DataFrame()

    def update_live_data(_self, creds_json: dict, sheet_url: str, risk_ids: List[str], new_status: str) -> bool:
        """Update risk status in Google Sheets with batch operations"""
        if not risk_ids:
            return True
            
        gc = _self._get_gspread_client(creds_json)
        if not gc:
            return False
            
        try:
            spreadsheet = gc.open_by_url(sheet_url)
            worksheet = spreadsheet.sheet1
            headers = worksheet.row_values(1)
            
            # Validate required headers
            try:
                risk_id_col = headers.index('Risk ID') + 1
                status_col = headers.index('Status') + 1
                last_updated_col = headers.index('Last Updated') + 1
            except ValueError as e:
                st.error("Google Sheet must contain headers: 'Risk ID', 'Status', 'Last Updated'.")
                return False

            today_str = date.today().strftime('%Y-%m-%d')
            updated_count = 0

            # Use batch updates for better performance
            updates = []
            for risk_id in risk_ids:
                # Find the row with this risk ID
                try:
                    cell = worksheet.find(risk_id, in_column=risk_id_col)
                    updates.extend([
                        {'range': rowcol_to_a1(cell.row, status_col), 'values': [[new_status]]},
                        {'range': rowcol_to_a1(cell.row, last_updated_col), 'values': [[today_str]]}
                    ])
                    updated_count += 1
                except gspread.exceptions.CellNotFound:
                    st.warning(f"Risk ID '{risk_id}' not found in Google Sheet")
                    continue

            if updates:
                worksheet.batch_update(updates)
                st.success(f"Updated {updated_count} risk(s) in Google Sheet")
                
            # Clear cache to reflect changes
            try:
                st.cache_data.clear()
            except:
                pass

            return updated_count > 0
            
        except Exception as e:
            st.error(f"Failed to update Google Sheet: {e}")
            return False

    def upload_df_to_gsheet(_self, creds_json: dict, sheet_url: str, df: pd.DataFrame) -> bool:
        """Upload DataFrame to Google Sheets with validation"""
        if df.empty:
            st.warning("No data to upload")
            return False
            
        gc = _self._get_gspread_client(creds_json)
        if not gc: 
            return False
            
        try:
            spreadsheet = gc.open_by_url(sheet_url)
            worksheet = spreadsheet.sheet1
            
            # Get existing headers to match columns
            headers = worksheet.row_values(1)
            upload_cols = [col for col in df.columns if col in headers]
            
            if not upload_cols:
                st.error("No matching columns found between DataFrame and Google Sheet")
                return False
                
            df_to_upload = df[upload_cols].copy()
            rows_to_append = df_to_upload.astype(str).values.tolist()
            
            worksheet.append_rows(rows_to_append, value_input_option='USER_ENTERED')
            
            # Clear cache
            try:
                st.cache_data.clear()
            except:
                pass
                
            return True
            
        except Exception as e:
            st.error(f"Failed to upload data to Google Sheet: {e}")
            return False

    @staticmethod
    def read_from_file(uploaded_file) -> pd.DataFrame:
        """Read data from uploaded file with better error handling"""
        if uploaded_file is None:
            return pd.DataFrame()
            
        try:
            if uploaded_file.name.lower().endswith('.csv'):
                return pd.read_csv(uploaded_file)
            else:
                return pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Error reading file {uploaded_file.name}: {e}")
            return pd.DataFrame()

    def preprocess_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and preprocess the risk data"""
        if df.empty:
            return df
            
        df_clean = df.copy()
        
        # Ensure required columns exist
        for col in Config.REQUIRED_COLUMNS:
            if col not in df_clean.columns:
                if col == 'Risk Score':
                    df_clean[col] = 1
                elif col in ['Likelihood', 'Impact']:
                    df_clean[col] = 1
                else:
                    df_clean[col] = 'Unknown'
        
        # Convert numeric columns
        for col in ['Likelihood', 'Impact', 'Risk Score']:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(1).astype(int)
        
        # Calculate risk score if not provided or invalid
        mask = (df_clean['Risk Score'] <= 0) | (df_clean['Risk Score'] > 25)
        df_clean.loc[mask, 'Risk Score'] = df_clean.loc[mask, 'Likelihood'] * df_clean.loc[mask, 'Impact']
        
        # Handle date column
        if 'Last Updated' in df_clean.columns:
            df_clean['Last Updated'] = pd.to_datetime(df_clean['Last Updated'], errors='coerce').dt.date
            df_clean['Last Updated'] = df_clean['Last Updated'].fillna(date.today())
        else:
            df_clean['Last Updated'] = date.today()
        
        # Fill missing categorical values
        defaults = {
            'Status': 'Open', 
            'Risk Owner': 'Unknown', 
            'Control Effectiveness': 'Medium',
            'Risk Category': 'Uncategorized'
        }
        
        for col, default in defaults.items():
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].fillna(default)
            else:
                df_clean[col] = default
        
        return df_clean

    def filter_data(self, df: pd.DataFrame, filters: Dict) -> pd.DataFrame:
        """Filter DataFrame based on user selections"""
        if df.empty:
            return df
            
        filtered_df = df.copy()
        
        # Apply filters
        if filters.get('owner') != 'All':
            filtered_df = filtered_df[filtered_df['Risk Owner'] == filters['owner']]
            
        if filters.get('status') != 'All':
            filtered_df = filtered_df[filtered_df['Status'] == filters['status']]
            
        if filters.get('control') != 'All':
            filtered_df = filtered_df[filtered_df['Control Effectiveness'] == filters['control']]
            
        # Score range filter
        score_range = filters.get('score_range', (1, 25))
        filtered_df = filtered_df[
            (filtered_df['Risk Score'] >= score_range[0]) & 
            (filtered_df['Risk Score'] <= score_range[1])
        ]
        
        return filtered_df.reset_index(drop=True)

# --- ENHANCED VISUALIZATION MANAGER ---
class VisualizationManager:
    @staticmethod
    def create_risk_matrix(df: pd.DataFrame) -> go.Figure:
        """Create an interactive risk matrix heatmap"""
        if df.empty:
            fig = go.Figure()
            fig.update_layout(
                title='Risk Matrix (No data)', 
                template='plotly_dark',
                annotations=[dict(
                    text="No data available",
                    xref="paper", yref="paper",
                    x=0.5, y=0.5, showarrow=False,
                    font=dict(size=20)
                )]
            )
            return fig
            
        try:
            # Create count matrix
            count_matrix = pd.pivot_table(
                df, 
                index='Impact', 
                columns='Likelihood', 
                aggfunc='size', 
                fill_value=0
            ).reindex(index=range(1, 6), columns=range(1, 6), fill_value=0)
            
            # Create risk level matrix for coloring
            risk_level_matrix = np.outer(range(1, 6), range(1, 6))
            
            colorscale = [
                [0.0, '#2E7D32'],   # Green - low risk
                [0.25, '#FFEB3B'],  # Yellow - medium risk
                [0.75, '#FF9800'],  # Orange - high risk
                [1.0, '#B71C1C']    # Red - critical risk
            ]
            
            fig = go.Figure(data=go.Heatmap(
                z=risk_level_matrix,
                x=count_matrix.columns,
                y=count_matrix.index,
                colorscale=colorscale,
                hoverinfo='text',
                hovertemplate=(
                    "<b>Impact:</b> %{y}<br>"
                    "<b>Likelihood:</b> %{x}<br>"
                    "<b>Risk Level:</b> %{z}<br>"
                    "<b>Number of Risks:</b> %{customdata}<extra></extra>"
                ),
                customdata=count_matrix.values,
                colorbar=dict(title="Risk Level")
            ))
            
            # Add annotations with risk counts
            annotations = []
            for i in count_matrix.index:
                for j in count_matrix.columns:
                    count_val = count_matrix.loc[i, j]
                    if count_val > 0:
                        annotations.append(dict(
                            x=j, y=i,
                            text=f"<b>{count_val}</b>",
                            showarrow=False,
                            font=dict(
                                color="white" if (i * j) >= 10 else "black",
                                size=16
                            )
                        ))
            
            fig.update_layout(
                template='plotly_dark',
                title=dict(
                    text="<b>Risk Heatmap</b> (Count of Risks)",
                    x=0.5,
                    font=dict(size=20)
                ),
                xaxis=dict(
                    title='<b>Likelihood</b>',
                    side="bottom",
                    tickmode='array',
                    tickvals=list(range(1, 6))
                ),
                yaxis=dict(
                    title='<b>Impact</b>',
                    tickmode='array',
                    tickvals=list(range(1, 6))
                ),
                height=500,
                annotations=annotations,
                margin=dict(l=60, r=40, t=80, b=60)
            )
            
            return fig
            
        except Exception as e:
            st.error(f"Error creating risk matrix: {e}")
            return go.Figure()

    @staticmethod
    def create_distribution_charts(df: pd.DataFrame) -> go.Figure:
        """Create distribution charts for risk analysis"""
        if df.empty:
            fig = go.Figure()
            fig.update_layout(
                title='No data for distribution', 
                template='plotly_dark'
            )
            return fig
            
        try:
            fig = make_subplots(
                rows=1, cols=3,
                subplot_titles=("By Risk Owner", "By Category", "By Status"),
                horizontal_spacing=0.1
            )
            
            # Risk Owner distribution
            if 'Risk Owner' in df.columns:
                owner_counts = df['Risk Owner'].value_counts().head(10)  # Top 10 only
                fig.add_trace(
                    go.Bar(
                        x=owner_counts.values, 
                        y=owner_counts.index,
                        orientation='h',
                        name='By Owner',
                        marker_color='#1f77b4'
                    ), 
                    row=1, col=1
                )
            
            # Risk Category distribution  
            if 'Risk Category' in df.columns:
                category_counts = df['Risk Category'].value_counts()
                fig.add_trace(
                    go.Bar(
                        x=category_counts.index,
                        y=category_counts.values,
                        name='By Category',
                        marker_color='#ff7f0e'
                    ),
                    row=1, col=2
                )
            
            # Status distribution
            if 'Status' in df.columns:
                status_counts = df['Status'].value_counts()
                fig.add_trace(
                    go.Bar(
                        x=status_counts.index,
                        y=status_counts.values,
                        name='By Status',
                        marker_color='#2ca02c'
                    ),
                    row=1, col=3
                )
            
            fig.update_layout(
                title_text="Risk Distribution Analysis",
                showlegend=False,
                template='plotly_dark',
                height=400,
                margin=dict(t=80)
            )
            
            # Update axes
            fig.update_xaxes(title_text="Count", row=1, col=1)
            fig.update_xaxes(title_text="Count", row=1, col=2)
            fig.update_xaxes(title_text="Count", row=1, col=3)
            fig.update_yaxes(title_text="Risk Owner", row=1, col=1)
            
            return fig
            
        except Exception as e:
            st.error(f"Error creating distribution charts: {e}")
            return go.Figure()

    @staticmethod
    def create_control_effectiveness_chart(df: pd.DataFrame) -> go.Figure:
        """Create control effectiveness analysis chart"""
        if df.empty or 'Control Effectiveness' not in df.columns:
            fig = go.Figure()
            fig.update_layout(
                title='No data for control analysis', 
                template='plotly_dark'
            )
            return fig
            
        try:
            # Calculate statistics
            effectiveness_stats = df.groupby('Control Effectiveness').agg({
                'Risk Score': ['count', 'mean', 'max']
            }).round(2)
            effectiveness_stats.columns = ['Count', 'Average Score', 'Max Score']
            effectiveness_stats = effectiveness_stats.reset_index()
            
            # Ensure proper ordering
            effectiveness_order = {"Low": 0, "Medium": 1, "High": 2}
            effectiveness_stats['Order'] = effectiveness_stats['Control Effectiveness'].map(effectiveness_order)
            effectiveness_stats = effectiveness_stats.sort_values('Order')
            
            fig = px.bar(
                effectiveness_stats, 
                x='Control Effectiveness', 
                y='Average Score',
                title='Average Risk Score by Control Effectiveness',
                color='Control Effectiveness',
                color_discrete_map={
                    'Low': Config.COLORS['high_risk'],
                    'Medium': Config.COLORS['medium_risk'],
                    'High': Config.COLORS['low_risk']
                },
                category_orders={"Control Effectiveness": ["Low", "Medium", "High"]}
            )
            
            # Add count annotations
            for i, row in effectiveness_stats.iterrows():
                fig.add_annotation(
                    x=row['Control Effectiveness'],
                    y=row['Average Score'] + 0.5,
                    text=f"n={row['Count']}",
                    showarrow=False,
                    font=dict(color='white', size=12)
                )
            
            fig.update_layout(
                template='plotly_dark',
                height=400,
                yaxis_title="Average Risk Score",
                showlegend=False
            )
            
            return fig
            
        except Exception as e:
            st.error(f"Error creating control effectiveness chart: {e}")
            return go.Figure()

# --- IMPROVED UI COMPONENTS ---
class UIComponents:
    @staticmethod
    def apply_custom_styling():
        """Apply custom CSS styling for better appearance"""
        st.markdown(f"""
            <style>
            .main .block-container {{
                padding-top: 2rem;
            }}
            .stApp {{
                background-color: {Config.COLORS['background']};
            }}
            .card {{
                background: {Config.COLORS['card_background']};
                border-radius: 10px;
                padding: 20px;
                margin: 10px 0;
                color: white;
                border-left: 5px solid;
                box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2);
            }}
            .kpi-title {{
                font-size: 14px;
                color: #e0e0e0;
                margin-bottom: 8px;
                font-weight: 500;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            .kpi-value {{
                font-size: 28px;
                font-weight: 700;
                margin: 0;
            }}
            .risk-high {{ border-left-color: {Config.COLORS['high_risk']}; }}
            .risk-medium {{ border-left-color: {Config.COLORS['medium_risk']}; }}
            .risk-low {{ border-left-color: {Config.COLORS['low_risk']}; }}
            .risk-critical {{ border-left-color: {Config.COLORS['critical_risk']}; }}
            
            /* Tab styling */
            .stTabs [data-baseweb="tab-list"] {{
                gap: 8px;
            }}
            .stTabs [data-baseweb="tab"] {{
                background-color: transparent;
                border-radius: 8px 8px 0px 0px;
                padding: 10px 16px;
            }}
            .stTabs [aria-selected="true"] {{
                background-color: {Config.COLORS['card_background']};
            }}
            
            /* Data editor improvements */
            .dataframe {{
                font-size: 14px;
            }}
            </style>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def render_kpi_card(title: str, value, risk_level: str = "low", icon: str = ""):
        """Render a KPI card with consistent styling"""
        st.markdown(
            f'<div class="card risk-{risk_level}">'
            f'<div class="kpi-title">{icon} {title}</div>'
            f'<div class="kpi-value">{value}</div>'
            f'</div>', 
            unsafe_allow_html=True
        )
    
    @staticmethod
    def get_risk_level(score: float) -> Tuple[str, str]:
        """Determine risk level based on score"""
        if score >= 20: 
            return "critical", Config.COLORS['critical_risk']
        elif score >= 15: 
            return "high", Config.COLORS['high_risk']
        elif score >= 8: 
            return "medium", Config.COLORS['medium_risk']
        else: 
            return "low", Config.COLORS['low_risk']

    @staticmethod
    def render_data_source_status(is_live: bool, df_loaded: bool):
        """Display data source status"""
        col1, col2 = st.columns([3, 1])
        with col1:
            if df_loaded:
                source_type = "Google Sheets (Live)" if is_live else "File Upload"
                st.success(f"✅ Data loaded from {source_type}")
            else:
                st.info("📊 Configure data source to begin")
        with col2:
            if df_loaded:
                st.button("🔄 Refresh", use_container_width=True)

# --- IMPROVED REPORT MANAGER ---
class ReportManager:
    class GRCReportTemplate(SimpleDocTemplate):
        def __init__(self, filename, **kw):
            super().__init__(filename, pagesize=landscape(letter), **kw)

        def afterPage(self):
            canvas = self.canv
            canvas.saveState()
            canvas.setFont('Helvetica', 9)
            canvas.drawString(inch, 0.75 * inch, f"Page {self.page} | GRC Risk Report")
            canvas.drawRightString(self.width + self.leftMargin - inch, 0.75 * inch, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            canvas.restoreState()

    @staticmethod
    def generate_excel_report(df: pd.DataFrame, risk_matrix_fig: go.Figure, filters: Dict, session_mitigated_df: pd.DataFrame) -> bytes:
        """Generate comprehensive Excel report"""
        if not OPENPYXL_AVAILABLE:
            st.error("openpyxl not available - cannot generate Excel reports")
            return b""
            
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Report information
                report_info = pd.DataFrame([
                    ["Report Generated", datetime.now().strftime('%Y-%m-%d %H:%M')],
                    ["Total Risks", len(df)],
                    ["Active Risks", len(df[~df['Status'].isin(['Mitigated', 'Closed'])])],
                    ["Mitigated Risks", len(df[df['Status'] == 'Mitigated'])]
                ] + [[f"Filter: {k}", str(v)] for k, v in filters.items()])
                
                report_info.to_excel(writer, index=False, header=False, sheet_name='Report_Info')
                
                # Main data
                df.to_excel(writer, index=False, sheet_name='All_Risks')
                
                # Filtered risks
                active_risks = df[~df['Status'].isin(['Mitigated', 'Closed'])]
                if not active_risks.empty:
                    active_risks.to_excel(writer, index=False, sheet_name='Active_Risks')
                
                # Mitigated risks
                mitigated_df = df[df['Status'] == 'Mitigated']
                if not mitigated_df.empty:
                    mitigated_df.to_excel(writer, index=False, sheet_name='All_Mitigated_Risks')
                
                # Session mitigated risks
                if not session_mitigated_df.empty:
                    session_mitigated_df.to_excel(writer, index=False, sheet_name='Mitigated_This_Session')

                # Risk matrix image
                try:
                    img_bytes = pio.to_image(risk_matrix_fig, format="png", width=800, height=600, scale=2)
                    workbook = writer.book
                    matrix_ws = workbook.create_sheet("Risk_Matrix")
                    img = OpenpyxlImage(io.BytesIO(img_bytes))
                    matrix_ws.add_image(img, 'A1')
                except Exception as e:
                    st.warning(f"Could not add risk matrix to Excel: {e}")

                # Auto-adjust column widths
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                        
            return output.getvalue()
            
        except Exception as e:
            st.error(f"Error generating Excel report: {e}")
            return b""

    @staticmethod
    def generate_pdf_report(df: pd.DataFrame, risk_matrix_fig: go.Figure, filters: Dict, session_mitigated_df: pd.DataFrame) -> bytes:
        """Generate comprehensive PDF report"""
        buffer = io.BytesIO()
        try:
            doc = ReportManager.GRCReportTemplate(buffer)
            styles = getSampleStyleSheet()
            elements = []
            
            # Title and metadata
            elements.append(Paragraph("GRC Risk Management Report", styles['Title']))
            elements.append(Spacer(1, 12))
            elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles['Normal']))
            elements.append(Paragraph(f"Total Risks: {len(df)} | Active: {len(df[~df['Status'].isin(['Mitigated', 'Closed'])])}", styles['Normal']))
            elements.append(Spacer(1, 12))
            
            # Filters section
            elements.append(Paragraph("<b>Active Filters</b>", styles['Heading2']))
            filter_text = "<br/>".join([f"<b>{k.replace('_', ' ').title()}:</b> {v}" for k, v in filters.items()])
            elements.append(Paragraph(filter_text, styles['Normal']))
            elements.append(PageBreak())
            
            # Risk matrix
            elements.append(Paragraph("<b>Risk Assessment Matrix</b>", styles['Heading2']))
            try:
                img_bytes = pio.to_image(risk_matrix_fig, format="png", width=550, height=412)
                elements.append(ReportLabImage(io.BytesIO(img_bytes)))
            except Exception as e:
                elements.append(Paragraph(f"Could not generate risk matrix: {e}", styles['Normal']))
            elements.append(PageBreak())

            # Table styling
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#DCE6F1")),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 10),
            ])
            
            report_cols = ['Risk ID', 'Title', 'Risk Owner', 'Risk Score', 'Status']

            # Active risks table
            non_mitigated_df = df[~df['Status'].isin(['Mitigated', 'Closed', 'Accepted'])].sort_values('Risk Score', ascending=False)
            if not non_mitigated_df.empty:
                elements.append(Paragraph("<b>Active (Non-Mitigated) Risks</b>", styles['Heading2']))
                data = [report_cols] + non_mitigated_df[report_cols].head(20).values.tolist()  # Limit to first 20
                table = Table(data, colWidths=[1*inch, 3.5*inch, 1.5*inch, 0.8*inch, 1*inch])
                table.setStyle(table_style)
                elements.append(table)
                if len(non_mitigated_df) > 20:
                    elements.append(Paragraph(f"... and {len(non_mitigated_df) - 20} more risks", styles['Normal']))
                elements.append(PageBreak())

            # Session mitigated risks
            if not session_mitigated_df.empty:
                elements.append(Paragraph("<b>Risks Mitigated This Session</b>", styles['Heading2']))
                data = [report_cols] + session_mitigated_df[report_cols].values.tolist()
                table = Table(data, colWidths=[1*inch, 3.5*inch, 1.5*inch, 0.8*inch, 1*inch])
                table.setStyle(table_style)
                elements.append(table)
                elements.append(Spacer(1, 24))

            # All mitigated risks
            mitigated_df = df[df['Status'] == 'Mitigated'].sort_values('Risk Score', ascending=False)
            if not mitigated_df.empty:
                elements.append(Paragraph("<b>All Mitigated Risks</b>", styles['Heading2']))
                data = [report_cols] + mitigated_df[report_cols].head(15).values.tolist()  # Limit to first 15
                table = Table(data, colWidths=[1*inch, 3.5*inch, 1.5*inch, 0.8*inch, 1*inch])
                table.setStyle(table_style)
                elements.append(table)
                if len(mitigated_df) > 15:
                    elements.append(Paragraph(f"... and {len(mitigated_df) - 15} more mitigated risks", styles['Normal']))

            doc.build(elements)
            return buffer.getvalue()
            
        except Exception as e:
            st.error(f"Error generating PDF report: {e}")
            return b""

# --- MAIN APPLICATION WITH IMPROVED ERROR HANDLING ---
def main():
    # Initialize session state
    if 'last_filters' not in st.session_state:
        st.session_state.last_filters = {}
    if 'df' not in st.session_state:
        st.session_state.df = pd.DataFrame()
    if 'initial_mitigated_ids' not in st.session_state:
        st.session_state.initial_mitigated_ids = set()
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False

    # Set up page
    st.title(f"{Config.PAGE_ICON} {Config.PAGE_TITLE}")
    UIComponents.apply_custom_styling()
    
    data_manager = DataManager()
    
    # Sidebar - Data Source Configuration
    with st.sidebar:
        st.header("🔧 Configuration")
        
        with st.expander("🔄 **Data Source**", expanded=True):
            data_source = st.radio(
                "Choose source:", 
                ("File Upload", "Google Sheets (Live)"), 
                key="data_source", 
                horizontal=True
            )
            
            if data_source == "Google Sheets (Live)":
                gsheet_url = st.text_input(
                    "Google Sheet URL", 
                    value=st.session_state.get('gsheet_url', ''),
                    help="Paste the URL of your Google Sheet"
                )
                gsheet_creds_file = st.file_uploader(
                    "Upload Credentials JSON", 
                    type=['json'],
                    help="Upload your Google Service Account credentials JSON file"
                )
                
                if gsheet_url:
                    st.session_state.gsheet_url = gsheet_url
                if gsheet_creds_file:
                    st.session_state.gsheet_creds_file = gsheet_creds_file
                    
            else:  # File Upload
                uploaded_file = st.file_uploader(
                    "Upload Risk Register", 
                    type=['csv', 'xlsx', 'xls'],
                    help="Upload your risk register as CSV or Excel file"
                )
                if uploaded_file:
                    st.session_state.uploaded_file = uploaded_file

    # Load data based on selected source
    df_loaded = False
    creds_info = None
    is_live = False
    
    try:
        if data_source == "Google Sheets (Live)":
            if (st.session_state.get('gsheet_url') and 
                st.session_state.get('gsheet_creds_file')):
                
                is_live = True
                st.session_state.gsheet_creds_file.seek(0)
                creds_info = json.load(st.session_state.gsheet_creds_file)
                
                with st.spinner("📥 Loading data from Google Sheets..."):
                    raw_df = data_manager.read_live_data(creds_info, st.session_state.gsheet_url)
                    
                if not raw_df.empty:
                    # Only update if data has changed
                    current_hash = pd.util.hash_pandas_object(raw_df).sum()
                    previous_hash = st.session_state.get('data_hash', 0)
                    
                    if current_hash != previous_hash:
                        st.session_state.raw_df = raw_df
                        st.session_state.df = data_manager.preprocess_data(raw_df)
                        st.session_state.data_hash = current_hash
                        st.session_state.initial_mitigated_ids = set(
                            st.session_state.df[st.session_state.df['Status'] == 'Mitigated']['Risk ID']
                        )
                    
                    df_loaded = True
                    st.session_state.data_loaded = True
                    
        else:  # File Upload
            if 'uploaded_file' in st.session_state and st.session_state.uploaded_file:
                is_live = False
                with st.spinner("📥 Loading data from file..."):
                    raw_df = data_manager.read_from_file(st.session_state.uploaded_file)
                    
                if not raw_df.empty:
                    current_hash = pd.util.hash_pandas_object(raw_df).sum()
                    previous_hash = st.session_state.get('data_hash', 0)
                    
                    if current_hash != previous_hash:
                        st.session_state.raw_df = raw_df
                        st.session_state.df = data_manager.preprocess_data(raw_df)
                        st.session_state.data_hash = current_hash
                        st.session_state.initial_mitigated_ids = set(
                            st.session_state.df[st.session_state.df['Status'] == 'Mitigated']['Risk ID']
                        )
                    
                    df_loaded = True
                    st.session_state.data_loaded = True
                    
    except Exception as e:
        st.error(f"Error loading data: {e}")
        df_loaded = False

    df = st.session_state.get('df', pd.DataFrame())
    
    # Display data source status
    UIComponents.render_data_source_status(is_live, df_loaded)

    # Sidebar - Add New Risk Form
    with st.sidebar:
        with st.expander("➕ **Add New Risk**", expanded=False):
            with st.form("new_risk_form", clear_on_submit=True):
                st.subheader("New Risk Details")
                
                title = st.text_input("Risk Title*", placeholder="Enter risk description")
                owner = st.selectbox("Risk Owner*", Config.RISK_OWNERS)
                category = st.selectbox("Risk Category*", Config.RISK_CATEGORIES)
                
                c1, c2 = st.columns(2)
                with c1:
                    likelihood = st.slider("Likelihood*", 1, 5, 3, 
                                         help="1 = Very Unlikely, 5 = Very Likely")
                with c2:
                    impact = st.slider("Impact*", 1, 5, 3,
                                     help="1 = Minimal Impact, 5 = Severe Impact")
                
                control = st.selectbox("Control Effectiveness*", 
                                     Config.CONTROL_EFFECTIVENESS_OPTIONS, 
                                     index=1)
                
                submitted = st.form_submit_button(
                    "Add Risk", 
                    type="primary", 
                    use_container_width=True,
                    disabled=not df_loaded
                )
                
                if submitted:
                    if not title:
                        st.error("Please provide a risk title")
                    elif not df_loaded:
                        st.error("Please load data before adding risks")
                    else:
                        new_risk_score = likelihood * impact
                        
                        # Generate new Risk ID
                        if not df.empty and 'Risk ID' in df.columns:
                            try:
                                # Extract numeric part from existing Risk IDs
                                numeric_ids = pd.to_numeric(
                                    df['Risk ID'].str.extract(r'R-(\d+)')[0], 
                                    errors='coerce'
                                ).dropna()
                                max_id = int(numeric_ids.max()) if not numeric_ids.empty else 999
                                new_id = f"R-{max_id + 1:04d}"
                            except:
                                new_id = f"R-{len(df) + 1:04d}"
                        else:
                            new_id = "R-0001"

                        new_row = pd.DataFrame([{
                            'Risk ID': new_id,
                            'Title': title,
                            'Risk Owner': owner,
                            'Risk Category': category,
                            'Likelihood': likelihood,
                            'Impact': impact,
                            'Risk Score': new_risk_score,
                            'Status': 'Open',
                            'Control Effectiveness': control,
                            'Last Updated': date.today()
                        }])
                        
                        try:
                            if is_live:
                                with st.spinner("Adding risk to Google Sheet..."):
                                    success = data_manager.upload_df_to_gsheet(
                                        creds_info, 
                                        st.session_state.gsheet_url, 
                                        new_row
                                    )
                                if success:
                                    st.success(f"Added risk {new_id} to Google Sheet!")
                                    time.sleep(1)
                                    st.rerun()
                                else:
                                    st.error("Failed to add risk to Google Sheet")
                            else:
                                st.session_state.df = pd.concat([
                                    df, 
                                    data_manager.preprocess_data(new_row)
                                ], ignore_index=True)
                                st.success(f"Added risk {new_id} to session data!")
                                time.sleep(1)
                                st.rerun()
                                
                        except Exception as e:
                            st.error(f"Error adding new risk: {e}")

    # Main Dashboard
    if not df_loaded:
        st.info("👋 **Welcome!** Please configure a data source in the sidebar to begin.")
        
        # Show sample data structure
        with st.expander("📋 Expected Data Structure"):
            st.write("""
            Your data should include these columns:
            - **Risk ID**: Unique identifier (e.g., R-001)
            - **Title**: Risk description
            - **Risk Owner**: Department/team responsible
            - **Risk Category**: Type of risk
            - **Likelihood**: 1-5 scale
            - **Impact**: 1-5 scale  
            - **Risk Score**: Likelihood × Impact (auto-calculated if missing)
            - **Status**: Open, In Progress, Mitigated, etc.
            - **Control Effectiveness**: Low, Medium, High
            - **Last Updated**: Date of last modification
            """)
            
            sample_data = pd.DataFrame({
                'Risk ID': ['R-001', 'R-002'],
                'Title': ['Sample Risk 1', 'Sample Risk 2'],
                'Risk Owner': ['IT', 'Security'],
                'Risk Category': ['Data Protection', 'Cybersecurity'],
                'Likelihood': [3, 4],
                'Impact': [4, 3],
                'Risk Score': [12, 12],
                'Status': ['Open', 'Mitigated'],
                'Control Effectiveness': ['Medium', 'High'],
                'Last Updated': [date.today(), date.today()]
            })
            st.dataframe(sample_data, use_container_width=True)
        
        return

    # Calculate KPIs
    total_risks = len(df)
    mitigated_risks = len(df[df['Status'] == 'Mitigated'])
    avg_score = df['Risk Score'].mean() if not df.empty else 0
    critical_risks = len(df[df['Risk Score'] >= 20])
    
    level, _ = UIComponents.get_risk_level(avg_score)
    
    # Display KPIs
    c1, c2, c3, c4 = st.columns(4)
    with c1: 
        UIComponents.render_kpi_card("Total Risks", total_risks, level, "🗂️")
    with c2: 
        UIComponents.render_kpi_card("Mitigated Risks", mitigated_risks, "low", "✅")
    with c3: 
        UIComponents.render_kpi_card("Avg. Score", f"{avg_score:.1f}", level, "📈")
    with c4: 
        UIComponents.render_kpi_card("Critical Risks", critical_risks, 
                                   "critical" if critical_risks > 0 else "low", "🚨")
    
    st.markdown("---")

    # Sidebar - Filters
    with st.sidebar:
        with st.expander("🔍 **Filters**", expanded=True):
            if not df.empty:
                owners = ['All'] + sorted(df['Risk Owner'].dropna().unique())
                statuses = ['All'] + sorted(df['Status'].dropna().unique())
                controls = ['All'] + sorted(df['Control Effectiveness'].dropna().unique())
                
                owner_sel = st.selectbox("Risk Owner", owners)
                status_sel = st.selectbox("Status", statuses)
                control_sel = st.selectbox("Control Effectiveness", controls)
                
                score_min = int(df['Risk Score'].min()) if not df.empty else 1
                score_max = int(df['Risk Score'].max()) if not df.empty else 25
                score_sel = st.slider(
                    "Risk Score Range", 
                    score_min, 
                    score_max, 
                    (score_min, score_max)
                )
                
                current_filters = {
                    'owner': owner_sel, 
                    'status': status_sel, 
                    'control': control_sel, 
                    'score_range': score_sel
                }
                
                # Clear cached reports if filters change
                if st.session_state.last_filters != current_filters:
                    st.session_state.pop('excel_report', None)
                    st.session_state.pop('pdf_report', None)
                    
                st.session_state.last_filters = current_filters
                
            else:
                st.info("No data available for filtering")
                current_filters = {}
        
        # Export for file-based data
        if not is_live and df_loaded:
            with st.expander("📤 **Export Data**"):
                csv_data = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "Download Updated CSV", 
                    data=csv_data,
                    file_name=f"grc_risks_{date.today()}.csv",
                    use_container_width=True
                )

    # Apply filters
    filtered_df = data_manager.filter_data(df, current_filters) if not df.empty else pd.DataFrame()

    # Main Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 **Risk Register**", 
        "📊 **Analytics**", 
        "✅ **Mitigation**", 
        "📄 **Reports**"
    ])

    with tab1:
        st.header("Risk Register")
        if not filtered_df.empty:
            st.dataframe(
                filtered_df, 
                use_container_width=True, 
                height=600,
                hide_index=True
            )
            
            # Summary statistics
            st.subheader("Summary")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Filtered Risks", len(filtered_df))
            with col2:
                avg_filtered_score = filtered_df['Risk Score'].mean()
                st.metric("Average Score", f"{avg_filtered_score:.1f}")
            with col3:
                high_risk_count = len(filtered_df[filtered_df['Risk Score'] >= 15])
                st.metric("High/Critical Risks", high_risk_count)
        else:
            st.info("No risks match the current filters")

    with tab2:
        st.header("Risk Analytics Dashboard")
        
        if not filtered_df.empty:
            # Risk Matrix
            st.subheader("Risk Assessment Matrix")
            risk_matrix_fig = VisualizationManager.create_risk_matrix(filtered_df)
            st.plotly_chart(risk_matrix_fig, use_container_width=True)
            
            # Distribution Charts
            st.subheader("Risk Distribution")
            dist_fig = VisualizationManager.create_distribution_charts(filtered_df)
            st.plotly_chart(dist_fig, use_container_width=True)
            
            # Control Effectiveness
            st.subheader("Control Effectiveness Analysis")
            control_fig = VisualizationManager.create_control_effectiveness_chart(filtered_df)
            st.plotly_chart(control_fig, use_container_width=True)
            
        else:
            st.info("No data available for analytics with current filters")

    with tab3:
        st.header("🔄 Risk Mitigation Checklist")
        
        if not filtered_df.empty:
            # Get active risks (not mitigated or closed)
            active_risks = filtered_df[
                ~filtered_df['Status'].isin(['Mitigated', 'Closed', 'Accepted'])
            ].copy()
            
            if not active_risks.empty:
                # Prepare data for editing
                display_df = active_risks[[
                    'Risk ID', 'Title', 'Risk Owner', 'Risk Score', 'Status'
                ]].copy()
                display_df['Mitigate'] = False
                display_df = display_df.reset_index(drop=True)
                
                # Data editor for mitigation selection
                st.subheader("Select Risks to Mitigate")
                
                edited_df = st.data_editor(
                    display_df,
                    column_config={
                        'Mitigate': st.column_config.CheckboxColumn(
                            "Select to Mitigate",
                            help="Check to mark this risk for mitigation"
                        ),
                        'Risk Score': st.column_config.ProgressColumn(
                            "Risk Score",
                            min_value=1,
                            max_value=25,
                            format="%d"
                        ),
                        'Title': st.column_config.TextColumn(
                            "Risk Title",
                            width="large"
                        )
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="mitigation_editor"
                )
                
                # Get selected risks
                selected_risks = edited_df[edited_df['Mitigate'] == True]
                selected_ids = selected_risks['Risk ID'].tolist()
                
                # Progress tracking
                total_mitigatable = len(df[~df['Status'].isin(['Closed', 'Accepted'])])
                already_mitigated = len(df[df['Status'] == 'Mitigated'])
                newly_selected = len(selected_ids)
                
                total_progress = (already_mitigated + newly_selected) / total_mitigatable if total_mitigatable > 0 else 0
                
                st.subheader("Mitigation Progress")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Selected for Mitigation", newly_selected)
                with col2:
                    st.metric("Already Mitigated", already_mitigated)
                with col3:
                    st.metric("Total Progress", f"{total_progress:.1%}")
                
                st.progress(total_progress)
                
                # Mitigation action
                if newly_selected > 0:
                    if st.button(
                        f"🚀 Update {newly_selected} Risk(s) to Mitigated", 
                        type="primary",
                        use_container_width=True
                    ):
                        try:
                            success = False
                            
                            if is_live:
                                with st.spinner("Updating Google Sheet..."):
                                    success = data_manager.update_live_data(
                                        creds_info, 
                                        st.session_state.gsheet_url, 
                                        selected_ids, 
                                        'Mitigated'
                                    )
                            else:
                                # Update session state
                                st.session_state.df.loc[
                                    st.session_state.df['Risk ID'].isin(selected_ids), 
                                    'Status'
                                ] = 'Mitigated'
                                st.session_state.df.loc[
                                    st.session_state.df['Risk ID'].isin(selected_ids), 
                                    'Last Updated'
                                ] = date.today()
                                success = True
                            
                            if success:
                                st.success(f"Successfully mitigated {newly_selected} risk(s)!")
                                time.sleep(2)
                                st.rerun()
                            else:
                                st.error("Failed to update risks. Please try again.")
                                
                        except Exception as e:
                            st.error(f"Error during mitigation: {e}")
                else:
                    st.info("Select risks above to mark them as mitigated")
                    
            else:
                st.success("🎉 No active risks to mitigate with current filters!")
                
        else:
            st.info("No risks available for mitigation with current filters")

    with tab4:
        st.header("📄 Reports & Exports")
        
        if not filtered_df.empty:
            # Calculate session-mitigated risks
            current_mitigated_ids = set(df[df['Status'] == 'Mitigated']['Risk ID'])
            session_mitigated_ids = current_mitigated_ids - st.session_state.initial_mitigated_ids
            session_mitigated_df = df[df['Risk ID'].isin(session_mitigated_ids)]
            
            risk_matrix_fig = VisualizationManager.create_risk_matrix(filtered_df)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Excel Report")
                if st.button("Generate Excel Report", use_container_width=True):
                    with st.spinner("Creating Excel report..."):
                        excel_report = ReportManager.generate_excel_report(
                            filtered_df, 
                            risk_matrix_fig, 
                            current_filters, 
                            session_mitigated_df
                        )
                        if excel_report:
                            st.session_state.excel_report = excel_report
                            st.success("Excel report generated successfully!")
                
                if 'excel_report' in st.session_state and st.session_state.excel_report:
                    st.download_button(
                        "📥 Download Excel Report",
                        data=st.session_state.excel_report,
                        file_name=f"GRC_Report_{date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            with col2:
                st.subheader("PDF Report")
                if st.button("Generate PDF Report", use_container_width=True):
                    with st.spinner("Creating PDF report..."):
                        pdf_report = ReportManager.generate_pdf_report(
                            filtered_df, 
                            risk_matrix_fig, 
                            current_filters, 
                            session_mitigated_df
                        )
                        if pdf_report:
                            st.session_state.pdf_report = pdf_report
                            st.success("PDF report generated successfully!")
                
                if 'pdf_report' in st.session_state and st.session_state.pdf_report:
                    st.download_button(
                        "📥 Download PDF Report",
                        data=st.session_state.pdf_report,
                        file_name=f"GRC_Report_{date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
            
            # Quick exports
            st.subheader("Quick Exports")
            qcol1, qcol2 = st.columns(2)
            
            with qcol1:
                # Export filtered data as CSV
                csv_data = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📊 Export Filtered Data (CSV)",
                    data=csv_data,
                    file_name=f"filtered_risks_{date.today()}.csv",
                    use_container_width=True
                )
            
            with qcol2:
                # Export active risks
                active_risks = filtered_df[~filtered_df['Status'].isin(['Mitigated', 'Closed'])]
                if not active_risks.empty:
                    active_csv = active_risks.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "⚠️ Export Active Risks (CSV)",
                        data=active_csv,
                        file_name=f"active_risks_{date.today()}.csv",
                        use_container_width=True
                    )
                    
        else:
            st.warning("No data available for reporting with current filters")

    # Footer
    st.markdown("---")
    st.caption("Developed by Amritesh Shrivastava | BMSCE College")

if __name__ == "__main__":
    main()

