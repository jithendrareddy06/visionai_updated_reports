import os
import tempfile
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, MSO_ANCHOR
from pptx.util import Inches, Pt
from pptx import Presentation
import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import base64
import re
import textwrap
from io import BytesIO
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import calendar
import shutil
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
from openpyxl import load_workbook
warnings.filterwarnings('ignore')
def extract_hyperlinks_from_excel(excel_file, column_name='Event Video Path'):
    """
    Extract actual hyperlinks from Excel file instead of just display text
    """
    try:
        # Handle both file paths and Streamlit uploaded file objects
        if hasattr(excel_file, 'read'):
            # It's a Streamlit uploaded file object
            excel_file.seek(0)  # Reset file pointer to beginning
            wb = load_workbook(excel_file, data_only=False)
        else:
            # It's a file path
            wb = load_workbook(excel_file, data_only=False)
        ws = wb.active
        # Find the column index for the video path column
        video_col_index = None
        header_row = 1
        for col_idx, cell in enumerate(ws[header_row], 1):
            if cell.value and str(cell.value).strip() == column_name:
                video_col_index = col_idx
                break
        
        if video_col_index is None:
            return {}
        
        # Extract hyperlinks from the video path column
        hyperlinks = {}
        
        for row_idx in range(2, ws.max_row + 1):  # Start from row 2 (skip header)
            cell = ws.cell(row=row_idx, column=video_col_index)
            
            if cell.hyperlink and cell.hyperlink.target:
                # Use the hyperlink target (actual URL)
                hyperlinks[row_idx - 2] = cell.hyperlink.target  # row_idx - 2 to match pandas 0-based indexing
            elif cell.value and str(cell.value).strip():
                # Fallback to cell value if no hyperlink but has text
                cell_value = str(cell.value).strip()
                hyperlinks[row_idx - 2] = cell_value
        
        return hyperlinks
    
    except Exception as e:
        st.error(f"❌ Error in hyperlink extraction: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return {}


def create_safe_temp_dir():
    """Create a temporary directory in the current working directory for AWS compatibility"""
    try:
        # First try the standard tempfile approach
        return tempfile.mkdtemp()
    except (OSError, PermissionError):
        # If that fails (like on AWS), create in current directory
        import uuid
        temp_name = f"temp_charts_{uuid.uuid4().hex[:8]}"
        temp_path = os.path.join(os.getcwd(), temp_name)
        os.makedirs(temp_path, exist_ok=True)
        return temp_path

# PowerPoint generation imports

st.set_page_config(layout="wide")
st.title("Consolidated Events Report")

# Configure seaborn and matplotlib styling for professional charts
sns.set_style("whitegrid")
sns.set_palette("husl")
plt.rcParams.update({
    'font.size': 12,
    'font.family': 'sans-serif',
    'axes.titlesize': 14,
    'axes.labelsize': 12,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'legend.fontsize': 11,
    'figure.titlesize': 16,
    'axes.grid': True,
    'grid.alpha': 0.3
})

# Professional color scheme
PROFESSIONAL_BLUE = '#2563EB'  # Single blue for all bar charts
PIE_COLORS = ['#2563EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6']  # Minimal colors for pie charts


def value_to_color(val, vmin, vmax):
    if vmax == vmin:
        return '#ffffff'
    ratio = (val - vmin) / (vmax - vmin)
    r = 255
    g = int(102 + (255 - 102) * (1 - ratio))
    b = int(102 + (255 - 102) * (1 - ratio))
    return f'#{r:02x}{g:02x}{b:02x}'


def df_to_heatmap_html(df, apply_heatmap=True):
    df = df.copy()
    numeric_cols = [
        col for col in df.select_dtypes(
            include=[
            np.number]).columns]

    # Exclude Total column and Grand Total row from scaling calculation
    df_for_scaling = df.copy()
    if 'Total' in df_for_scaling.columns and 'Total' in numeric_cols:
        numeric_cols_for_scaling = [
            col for col in numeric_cols if col != 'Total']
    else:
        numeric_cols_for_scaling = numeric_cols

    if 'Grand Total' in df_for_scaling.index:
        df_for_scaling = df_for_scaling.drop('Grand Total', axis=0)

    # Calculate min/max across entire table (excluding totals)
    if numeric_cols_for_scaling and not df_for_scaling.empty:
        vmin = df_for_scaling[numeric_cols_for_scaling].min().min()
        vmax = df_for_scaling[numeric_cols_for_scaling].max().max()
    else:
        vmin = 0
        vmax = 1
    html = '<table class="main-table">'
    html += '<thead><tr>'
    html += '<th>Camera</th>'
    for col in df.columns:
        html += f'<th style="border: 1px solid #dee2e6; padding: 8px; text-align: center; font-weight: bold;">{col}</th>'
    html += '</tr></thead><tbody>'
    for idx, row in df.iterrows():
        html += f'<tr><td style="border: 1px solid #dee2e6; padding: 8px; font-weight: bold;">{idx}</td>'
        for col in df.columns:
            value = row[col]
            if col in numeric_cols and apply_heatmap:
                # Don't apply heatmap to Total columns or Grand Total rows
                if col == 'Total' or idx == 'Grand Total':
                    bg_color = '#f8f9fa'
                    text_color = '#000'
                    html += f'<td style="background-color:{bg_color};color:{text_color};font-weight:bold;">{value}</td>'
                else:
                    # Calculate color using entire table scaling
                    if vmax > vmin and value > 0:
                        intensity = (value - vmin) / (vmax - vmin)
                    else:
                        intensity = 0
                    # Color scale from white (0) to red (max)
                    if value == 0:
                        bg_color = '#ffffff'
                        text_color = '#6c757d'
                    else:
                        # Red intensity based on value
                        red = 255
                        green = int(255 - (intensity * 153))  # 255 to 102
                        blue = int(255 - (intensity * 153))   # 255 to 102
                        bg_color = f'rgb({red}, {green}, {blue})'
                        text_color = '#000000' if intensity < 0.7 else '#ffffff'
                    html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center; background-color: {bg_color}; color: {text_color}; font-weight: bold;">{value}</td>'
            else:
                html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center;">{value}</td>'
        html += '</tr>'

    html += '</tbody></table>'
    return html


def process_event_report(excel_file):
    try:
        df = pd.read_excel(
            excel_file,
            usecols=[
                'Camera',
                'Events',
                'Area',
                'Severity',
                'Reviewed'])
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None

    if df.empty:
        st.error("The Excel file contains no data.")
        return None

    required_columns = ['Camera', 'Events', 'Area', 'Severity', 'Reviewed']
    missing_columns = [
        col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing required columns: {', '.join(missing_columns)}")
        return None

    df = df.dropna(subset=['Camera', 'Events', 'Area', 'Severity'])  # Don't require Reviewed to be non-null
    if df.empty:
        st.error("No valid data found after removing rows with missing values.")
        return None

    valid_severities = ['Critical', 'High', 'Moderate', 'Low']
    invalid_severities = df[~df['Severity'].isin(
        valid_severities)]['Severity'].unique()
    if len(invalid_severities) > 0:
        st.error(
            f"Invalid severity values found: {', '.join(invalid_severities)}")
        return None

    area_totals = df.groupby('Area').size().sort_values(ascending=False)
    sorted_areas = area_totals.index.tolist()
    total_events_all = len(df)
    near_miss_df = df[df['Events'] == 'Near Miss']
    total_near_miss = len(near_miss_df)
    near_miss_by_area = near_miss_df.groupby('Area').size()
    near_miss_by_camera = near_miss_df.groupby(
        'Camera').size().sort_values(ascending=False)

    # Consolidated tables
    event_pivot_all = df.groupby(
        ['Camera', 'Events']).size().unstack(fill_value=0)
    event_pivot_all = event_pivot_all.loc[:, event_pivot_all.sum(axis=0) != 0]
    top_events_all = event_pivot_all.sum().sort_values(ascending=False).head(10).index
    event_pivot_all = event_pivot_all[top_events_all]
    event_pivot_all['Total'] = event_pivot_all.sum(axis=1)
    event_pivot_all = event_pivot_all.sort_values(
        by='Total', ascending=False).head(10)
    event_pivot_all.index.name = None
    event_pivot_all.columns.name = None

    camera_contributions = []
    for camera in event_pivot_all.index:
        camera_total = event_pivot_all.loc[camera, 'Total']
        contribution = (camera_total / total_events_all) * 100
        camera_contributions.append({
            'camera': camera,
            'total': camera_total,
            'contribution': contribution
        })
    camera_contributions = sorted(
        camera_contributions,
        key=lambda x: x['total'],
        reverse=True)[
        :3]

    grand_total_all = event_pivot_all.sum(axis=0)
    grand_total_all.name = 'Grand Total'
    event_pivot_all = pd.concat([event_pivot_all,
                                 pd.DataFrame([grand_total_all],
                                              columns=event_pivot_all.columns,
                                              index=[grand_total_all.name])])
    cols_to_keep = [col for col in event_pivot_all.columns if (
        col == 'Total' or event_pivot_all.loc['Grand Total', col] != 0)]
    event_pivot_all = event_pivot_all[cols_to_keep]

    severity_pivot_all = df.groupby(
        ['Camera', 'Severity']).size().unstack(fill_value=0)
    for level in ['Critical', 'High', 'Moderate', 'Low']:
        if level not in severity_pivot_all.columns:
            severity_pivot_all[level] = 0
    severity_pivot_all = severity_pivot_all[[
        'Critical', 'High', 'Moderate', 'Low']]
    severity_pivot_all['Total'] = severity_pivot_all.sum(axis=1)
    severity_pivot_all = severity_pivot_all.sort_values(
        by='Total', ascending=False).head(10)
    severity_pivot_all.index.name = None
    severity_pivot_all.columns.name = None
    grand_total_sev_all = severity_pivot_all.sum(axis=0)
    grand_total_sev_all.name = 'Grand Total'
    severity_pivot_all = pd.concat([severity_pivot_all,
                                    pd.DataFrame([grand_total_sev_all],
                                                 columns=severity_pivot_all.columns,
                                                 index=[grand_total_sev_all.name])])

    areas_data = []
    for area in sorted_areas:
        area_df = df[df['Area'] == area]
        total_events = len(area_df)
        area_contribution = (total_events / total_events_all) * 100
        area_near_miss = len(area_df[area_df['Events'] == 'Near Miss'])
        area_near_miss_contribution = (
            area_near_miss /
            total_near_miss *
            100) if total_near_miss > 0 else 0

        event_pivot = area_df.groupby(
            ['Camera', 'Events']).size().unstack(fill_value=0)
        event_pivot = event_pivot.loc[:, (event_pivot != 0).any(axis=0)]
        top_events = event_pivot.sum().sort_values(ascending=False).head(6).index
        event_pivot = event_pivot[top_events]
        event_pivot['Total'] = event_pivot.sum(axis=1)
        event_pivot = event_pivot.sort_values(
            by='Total', ascending=False).drop(
            columns='Total')
        event_pivot.index.name = None
        event_pivot.columns.name = None
        event_pivot['Total'] = event_pivot.sum(axis=1)
        grand_total = event_pivot.sum(axis=0)
        grand_total.name = 'Grand Total'
        event_pivot = pd.concat([event_pivot, pd.DataFrame(
            [grand_total], columns=event_pivot.columns, index=[grand_total.name])])

        severity_pivot = area_df.groupby(
            ['Camera', 'Severity']).size().unstack(fill_value=0)
        for level in ['Critical', 'High', 'Moderate', 'Low']:
            if level not in severity_pivot.columns:
                severity_pivot[level] = 0
        severity_pivot = severity_pivot[[
            'Critical', 'High', 'Moderate', 'Low']]
        severity_pivot['Total'] = severity_pivot.sum(axis=1)
        severity_pivot = severity_pivot.sort_values(
            by='Total', ascending=False).drop(
            columns='Total')
        severity_pivot.index.name = None
        severity_pivot.columns.name = None
        severity_pivot['Total'] = severity_pivot.sum(axis=1)
        grand_total_sev = severity_pivot.sum(axis=0)
        grand_total_sev.name = 'Grand Total'
        severity_pivot = pd.concat([severity_pivot,
                                    pd.DataFrame([grand_total_sev],
                                                 columns=severity_pivot.columns,
                                                 index=[grand_total_sev.name])])

        areas_data.append({
            'name': area,
            'total_events': total_events,
            'contribution': area_contribution,
            'event_table': event_pivot,
            'severity_table': severity_pivot,
            'near_miss_count': area_near_miss,
            'near_miss_contribution': area_near_miss_contribution
        })

    camera_to_area = dict(df[['Camera', 'Area']].drop_duplicates().values)

    return {
        'areas': areas_data,
        'consolidated_event_table': event_pivot_all,
        'consolidated_severity_table': severity_pivot_all,
        'camera_contributions': camera_contributions,
        'total_events_all': total_events_all,
        'total_near_miss': total_near_miss,
        'near_miss_by_camera': near_miss_by_camera.to_dict(),
        'near_miss_by_area': near_miss_by_area.to_dict(),
        'camera_to_area': camera_to_area,
    }


def plot_events_per_day(df, month, month_name, year):
    # Validate data before processing
    if df.empty:
        return ""

    # Group by date and count events, only keep non-zero counts
    events_per_day = df.groupby(df['Date'].dt.date)['Events'].count().reset_index()
    events_per_day.columns = ['Date', 'Event_Count']
    non_zero_events = events_per_day[events_per_day['Event_Count'] > 0]

    # Plot
    fig, ax = plt.subplots(figsize=(12, 5))

    if len(non_zero_events) > 0:
        # Sort events by date to ensure proper line connection
        non_zero_events = non_zero_events.sort_values('Date')

        # Plot line connecting non-zero events
        ax.plot(
            non_zero_events['Date'],
            non_zero_events['Event_Count'],
            linewidth=2,
            color='purple',
            alpha=0.7)

        # Add scatter points for non-zero events
        ax.scatter(
            non_zero_events['Date'],
            non_zero_events['Event_Count'],
            color='purple',
            s=60,
            zorder=5,
            label='With Events')

        # Add value labels for non-zero points
        for date, count in zip(non_zero_events['Date'], non_zero_events['Event_Count']):
            ax.annotate(str(int(count)),
                        (date, count),
                        textcoords="offset points",
                        xytext=(0, 10),
                        ha='center',
                        fontsize=10,
                        fontweight='bold')

        # Set x-axis ticks to only show dates with events
        ax.set_xticks(non_zero_events['Date'])
        ax.set_xticklabels([d.strftime('%d') for d in non_zero_events['Date']], rotation=45)

        ax.legend()

    ax.set_title(f'Events Per Day - {month_name} {year}', fontsize=14, fontweight='bold')
    ax.set_xlabel('Date')
    ax.set_ylabel('Number of Events')
    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_avg_interval(df, month, month_name, year):
    # Convert Time column if it exists and has valid data
    if 'Time' in df.columns and not df['Time'].isna().all():
        df['Time'] = pd.to_datetime(df['Time'], dayfirst=True, errors='coerce')
        # Remove rows with invalid time data for interval calculation
        df_time_filtered = df.dropna(subset=['Time'])
    else:
        df_time_filtered = pd.DataFrame()  # Empty dataframe if no valid time data

    daily_intervals = []
    for date in df['Date'].dt.date.unique():
        if not df_time_filtered.empty:
            day_data = df_time_filtered[df_time_filtered['Date'].dt.date == date].sort_values(
                'Time')
        if len(day_data) > 1:
            time_diffs = [
                (day_data.iloc[i]['Time'] - day_data.iloc[i - 1]
                 ['Time']).total_seconds() / 60
                for i in range(1, len(day_data))
            ]
            avg_interval = sum(time_diffs) / len(time_diffs)
            daily_intervals.append(
                {'Date': date, 'Avg_Interval_Minutes': avg_interval})
        else:
            daily_intervals.append({'Date': date, 'Avg_Interval_Minutes': 0})

    intervals_df = pd.DataFrame(daily_intervals)
    # Create complete date range
    start_date = pd.Timestamp(year=year, month=month, day=1).date()
    if month in [4, 6, 9, 11]:
        end_date = pd.Timestamp(year=year, month=month, day=30).date()
    elif month == 2:
        if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
            end_date = pd.Timestamp(year=year, month=month, day=29).date()
        else:
            end_date = pd.Timestamp(year=year, month=month, day=28).date()
    else:
        end_date = pd.Timestamp(year=year, month=month, day=31).date()
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    complete_dates = pd.DataFrame({'Date': date_range.date})
    complete_intervals = complete_dates.merge(
        intervals_df, on='Date', how='left')
    complete_intervals['Avg_Interval_Minutes'] = complete_intervals['Avg_Interval_Minutes'].fillna(
        0)
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(
        complete_intervals['Date'],
        complete_intervals['Avg_Interval_Minutes'],
        linewidth=2,
        color="teal",
        alpha=0.7)
    zero_dates = complete_intervals[complete_intervals['Avg_Interval_Minutes'] == 0]['Date']
    zero_intervals = complete_intervals[complete_intervals['Avg_Interval_Minutes']
                                        == 0]['Avg_Interval_Minutes']
    non_zero_dates = complete_intervals[complete_intervals['Avg_Interval_Minutes'] > 0]['Date']
    non_zero_intervals = complete_intervals[complete_intervals['Avg_Interval_Minutes']
                                            > 0]['Avg_Interval_Minutes']
    if len(zero_dates) > 0:
        ax.scatter(
            zero_dates,
            zero_intervals,
            color='red',
            s=60,
            zorder=5,
            label='No Events/Single Event')
    if len(non_zero_dates) > 0:
        ax.scatter(
            non_zero_dates,
            non_zero_intervals,
            color="teal",
            s=60,
            zorder=5,
            label='Multiple Events')
    ax.set_title(
        f'Average Time Interval Between Events - {month_name} {year}',
        fontsize=14,
        fontweight='bold')
    ax.set_xlabel('Date')
    ax.set_ylabel('Average Interval (Minutes)')
    ax.legend()
    ax.set_xticks(complete_intervals['Date'])
    ax.set_xticklabels([d.strftime('%d')
                       for d in complete_intervals['Date']], rotation=45)
    for date, interval in zip(
            complete_intervals['Date'], complete_intervals['Avg_Interval_Minutes']):
        label = f'{interval:.1f}m' if interval > 0 else '0m'
        ax.annotate(
            label, (date, interval), textcoords="offset points", xytext=(
                0, 10), ha='center', fontsize=8)
    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_events_per_area_pie(df, max_labels=8):
    # Use Area column if available, otherwise use Camera column
    if 'Area' in df.columns:
        area_counts = df['Area'].value_counts()
    elif 'Camera' in df.columns:
        area_counts = df['Camera'].value_counts()
    else:
        # Return empty chart if no area/camera data
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.text(0.5, 0.5, 'No Area/Camera data available',
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12)
        ax.set_title(
            'Number of Events per Area',
            fontsize=14,
            fontweight='bold')
        plt.tight_layout()
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
        plt.close(fig)
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode("utf-8")
        return img_base64

    # Group small slices into 'Other'
    if len(area_counts) > max_labels:
        top_areas = area_counts[:max_labels]
        other_sum = area_counts[max_labels:].sum()
        area_counts = pd.concat([top_areas, pd.Series({'Other': other_sum})])

    # Create figure with two subplots side by side
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))

    # Create clean pie chart without labels
    colors = plt.cm.Set3(np.linspace(0, 1, len(area_counts)))
    wedges, texts, autotexts = ax1.pie(
        area_counts.values,
        labels=None,
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        textprops={'fontsize': 10, 'fontweight': 'bold'}
    )

    ax1.set_title('Number of Events per Area', fontsize=14, fontweight='bold')

    # Create legend table on the right
    ax2.axis('off')  # Hide axes for the legend

    total_events = area_counts.sum()

    # Create table data
    table_data = []
    for i, (area, count) in enumerate(area_counts.items()):
        percentage = (count / total_events * 100) if total_events > 0 else 0
        table_data.append([
            f"● {area}",  # Color indicator
            f"{count}",
            f"{percentage:.1f}%"
        ])

    # Create table
    table = ax2.table(
        cellText=table_data,
        colLabels=['Area', 'Count', 'Percentage'],
        cellLoc='left',
        loc='center',
        colWidths=[0.5, 0.2, 0.3]
    )

    # Style the table
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 2)

    # Color the first column cells to match pie slices
    for i in range(len(area_counts)):
        table[(i + 1, 0)].set_facecolor(colors[i])
        table[(i + 1, 0)].set_text_props(weight='bold', color='black')
        table[(i + 1, 1)].set_text_props(weight='bold')
        table[(i + 1, 2)].set_text_props(weight='bold')

    # Style header row
    for j in range(3):
        table[(0, j)].set_facecolor('#40466e')
        table[(0, j)].set_text_props(weight='bold', color='white')

    ax2.set_title('Area Breakdown', fontsize=14, fontweight='bold', pad=20)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_events_by_type_bar(df, event_types=None):
    # If event_types is provided, use that order and only those types
    if event_types is not None:
        event_counts = df['Events'].value_counts()
        event_counts = event_counts.reindex(event_types, fill_value=0)
    else:
        # Get top 10 event types by count
        event_counts = df['Events'].value_counts().head(10)

    total_events = event_counts.sum()

    # Create a larger figure to accommodate wrapped labels
    fig, ax = plt.subplots(figsize=(12, 6))
    bars = ax.bar(
        event_counts.index,
        event_counts.values,
        color=PROFESSIONAL_BLUE,
        alpha=0.8)
    ax.set_title('Events by Type', fontsize=16, fontweight='bold')
    ax.set_xlabel('Event Type', fontsize=12, fontweight='bold')
    ax.set_ylabel('Number of Events', fontsize=12, fontweight='bold')

    # Wrap long labels and set x-axis labels with improved formatting
    wrapped_labels = []
    for label in event_counts.index:
        # Wrap text if longer than 15 characters
        if len(label) > 15:
            wrapped_label = '\n'.join(textwrap.wrap(label, width=15))
        else:
            wrapped_label = label
        wrapped_labels.append(wrapped_label)

    # Set x-axis ticks and labels with better spacing and rotation
    ax.set_xticks(range(len(event_counts)))
    ax.set_xticklabels(wrapped_labels, rotation=45, ha='right', fontsize=9,
                       va='top', linespacing=0.8)

    # Adjust tick parameters for better spacing
    ax.tick_params(axis='x', which='major', pad=10)
    ax.tick_params(axis='y', which='major', labelsize=10)

    # Add value labels on top of bars with percentage
    for bar, count in zip(bars, event_counts.values):
        height = bar.get_height()
        percentage = (count / total_events * 100) if total_events > 0 else 0
        ax.annotate(f'{int(height)}\n({percentage:.1f}%)',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=9, fontweight='bold')

    # Adjust layout with extra bottom margin for wrapped labels
    plt.subplots_adjust(bottom=0.25)
    plt.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_events_by_area_bar(df):
    if 'Area' in df.columns:
        area_counts = df['Area'].value_counts().head(10)  # Limit to top 10 areas
        xlabel_text = 'Area (Top 10)'
        title_suffix = 'Area (Top 10)'
    elif 'Camera' in df.columns:
        area_counts = df['Camera'].value_counts().head(10)  # Limit to top 10 cameras
        xlabel_text = 'Camera (Top 10)'
        title_suffix = 'Camera (Top 10)'
    else:
        # Return empty image if no area/camera data
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.text(0.5, 0.5, 'No Area/Camera data available',
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=14)
        ax.set_title('Events by Area', fontsize=16, fontweight='bold')
        plt.tight_layout()
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode("utf-8")
        return img_base64
    total_events = area_counts.sum()
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(
        area_counts.index,
        area_counts.values,
        color=PROFESSIONAL_BLUE,
        alpha=0.8)
    ax.set_title(f'Events by {title_suffix}', fontsize=16, fontweight='bold')
    ax.set_xlabel(xlabel_text, fontsize=12, fontweight='bold')
    ax.set_ylabel('Event Count', fontsize=12, fontweight='bold')
    plt.xticks(rotation=45, ha='right', fontsize=10)
    plt.yticks(fontsize=10)
    # Add value labels on top of bars with percentage
    for bar, count in zip(bars, area_counts.values):
        height = bar.get_height()
        percentage = (count / total_events * 100) if total_events > 0 else 0
        ax.annotate(f'{int(height)}\n({percentage:.1f}%)',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=9, fontweight='bold')
    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_events_by_severity_pie(df):
    if 'Severity' not in df.columns:
        # Return empty pie chart if no severity data
        fig, ax = plt.subplots(figsize=(6, 5))
        ax.text(0.5, 0.5, 'No Severity data available',
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12)
        ax.set_title('Events by Severity', fontsize=14, fontweight='bold')
        plt.tight_layout()
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
        plt.close(fig)
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode("utf-8")
        return img_base64

    severity_counts = df['Severity'].value_counts()
    total_events = severity_counts.sum()

    # Create figure with two subplots side by side
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 6))

    # Define severity colors (red to green gradient)
    severity_colors = {
        'Critical': '#e74c3c',
        'High': '#f39c12',
        'Moderate': '#f1c40f',
        'Low': '#2ecc71'
    }

    # Get colors for actual severity levels in data
    colors = []
    for severity in severity_counts.index:
        if severity in severity_colors:
            colors.append(severity_colors[severity])
        else:
            # Use professional colors for unknown severity levels
            idx = len(colors) % len(PIE_COLORS)
            colors.append(PIE_COLORS[idx])

    # Create clean pie chart on the left
    wedges, texts, autotexts = ax1.pie(
        severity_counts.values,
        labels=None,  # No labels on pie chart
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        textprops={'fontsize': 11, 'fontweight': 'bold'}
    )

    ax1.set_title('Events by Severity', fontsize=14, fontweight='bold')

    # Create summary table on the right
    ax2.axis('off')  # Hide axes for the table

    # Create table data
    table_data = []
    for i, (severity, count) in enumerate(severity_counts.items()):
        percentage = (count / total_events * 100) if total_events > 0 else 0
        table_data.append([
            f"● {severity}",  # Color indicator
            f"{count}",
            f"{percentage:.1f}%"
        ])

    # Create table
    table = ax2.table(
        cellText=table_data,
        colLabels=['Severity Level', 'Count', 'Percentage'],
        cellLoc='left',
        loc='center',
        colWidths=[0.5, 0.25, 0.25]
    )

    # Style the table
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 2)

    # Color the first column cells to match pie slices
    for i, (severity, count) in enumerate(severity_counts.items()):
        table[(i + 1, 0)].set_facecolor(colors[i])
        table[(i + 1, 0)].set_text_props(weight='bold', color='white')
        table[(i + 1, 1)].set_text_props(weight='bold')
        table[(i + 1, 2)].set_text_props(weight='bold')

    # Style header row
    for j in range(3):
        table[(0, j)].set_facecolor('#34495e')
        table[(0, j)].set_text_props(weight='bold', color='white')

    ax2.set_title('Severity Breakdown', fontsize=14, fontweight='bold', pad=20)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_events_by_type_pie(df):
    if 'Events' not in df.columns:
        # Return empty chart if no events data
        fig, ax = plt.subplots(figsize=(6, 4))
        ax.text(0.5, 0.5, 'No Events data available',
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12)
        ax.set_title('Events by Type', fontsize=14, fontweight='bold')
        plt.tight_layout()
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
        plt.close(fig)
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode("utf-8")
        return img_base64

    # Get ALL event types without any filtering
    event_counts = df['Events'].value_counts()

    # Create figure with two subplots side by side - make it wider for more event types
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 8))

    # Create clean pie chart without labels
    # Generate colors dynamically for any number of event types
    import matplotlib.cm as cm
    if len(event_counts) <= 10:
        # Use predefined colors for small number of events
        colors = PIE_COLORS[:len(event_counts)]
        if len(event_counts) > len(PIE_COLORS):
            colors.extend(['#94A3B8'] * (len(event_counts) - len(PIE_COLORS)))
    else:
        # Use minimal colors for larger number of events
        colors = PIE_COLORS[:len(event_counts)]
        if len(event_counts) > len(PIE_COLORS):
            colors.extend(['#94A3B8'] * (len(event_counts) - len(PIE_COLORS)))

    # Create a counter to track which slice we're on
    slice_counter = [0]  # Use list to make it mutable in nested function

    def custom_autopct(pct):
        # Show percentage only for first 5 slices (top 5 events)
        current_slice = slice_counter[0]
        slice_counter[0] += 1

        if current_slice < 5:  # Top 5 events
            return f'{pct:.1f}%'
        else:  # All other events
            return ''

    wedges, texts, autotexts = ax1.pie(
        event_counts.values,
        labels=None,
        autopct=custom_autopct,
        startangle=140,
        colors=colors,
        textprops={'fontsize': 10, 'fontweight': 'bold'}
    )

    ax1.set_title('Events by Type', fontsize=14, fontweight='bold')

    # Create legend table on the right
    ax2.axis('off')  # Hide axes for the legend

    total_events = event_counts.sum()

    # Get only top 10 event types for the table
    top_10_events = event_counts.head(10)

    # Create table data - for top 10 events
    table_data = []
    for i, (event_type, count) in enumerate(top_10_events.items()):
        percentage = (count / total_events * 100) if total_events > 0 else 0
        table_data.append([
            f"● {event_type}",  # Color indicator
            f"{count}",
            f"{percentage:.1f}%"
        ])

    # Add "Others" row if there are more than 10 event types
    if len(event_counts) > 10:
        remaining_events = event_counts.iloc[10:]  # Events beyond top 10
        others_count = remaining_events.sum()
        others_percentage = (others_count / total_events * 100) if total_events > 0 else 0
        table_data.append([
            f"● Others",  # Color indicator
            f"{others_count}",
            f"{others_percentage:.1f}%"
        ])

    # Create table
    table = ax2.table(
        cellText=table_data,
        colLabels=['Event Type', 'Count', 'Percentage'],
        cellLoc='left',
        loc='center',
        colWidths=[0.5, 0.2, 0.3]
    )

    # Style the table - adjust font size based on number of event types in table (max 10)
    table.auto_set_font_size(False)
    table.set_fontsize(10)  # Normal font for top 10 events
    table.scale(1, 2)

    # Color the first column cells to match pie slices (for top 10 + Others if
    # applicable)
    for i in range(len(top_10_events)):
        table[(i + 1, 0)].set_facecolor(colors[i])
        table[(i + 1, 0)].set_text_props(weight='bold', color='black')
        table[(i + 1, 1)].set_text_props(weight='bold')
        table[(i + 1, 2)].set_text_props(weight='bold')

    # Add color for "Others" row if it exists
    if len(event_counts) > 10:
        others_row_index = len(top_10_events) + 1  # +1 because of header row
        # Use a gray color for "Others" row
        table[(others_row_index, 0)].set_facecolor('#999999')
        table[(others_row_index, 0)].set_text_props(
            weight='bold', color='white')
        table[(others_row_index, 1)].set_text_props(weight='bold')
        table[(others_row_index, 2)].set_text_props(weight='bold')

    # Style header row
    for j in range(3):
        table[(0, j)].set_facecolor('#40466e')
        table[(0, j)].set_text_props(weight='bold', color='white')

    ax2.set_title(
        'Event Type Breakdown',
        fontsize=14,
        fontweight='bold',
        pad=20)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_camera_events_pie(area_data, area_name, area_column):
    """Generate pie chart for Camera vs Events for a specific area"""
    if area_column == 'Camera':
        # Single camera breakdown by events
        if 'Events' in area_data.columns:
            event_counts = area_data['Events'].value_counts()
            title = f'Events Distribution - {area_name}'
        else:
            return ""
    else:
        # Multiple cameras breakdown - show camera distribution
        if 'Camera' in area_data.columns:
            event_counts = area_data['Camera'].value_counts()
            title = f'Camera Distribution - {area_name}'
        else:
            return ""

    if event_counts.empty:
        return ""

    # Create pie chart
    fig, ax = plt.subplots(figsize=(6, 5))
    colors = [
        '#FF6B6B',
        '#4ECDC4',
        '#45B7D1',
        '#96CEB4',
        '#FECA57',
        '#FF9FF3',
        '#54A0FF',
        '#5F27CD']
    colors = colors[:len(event_counts)]

    wedges, texts, autotexts = ax.pie(
        event_counts.values,
        labels=None,
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        textprops={'fontsize': 9, 'fontweight': 'bold'}
    )

    ax.set_title(title, fontsize=12, fontweight='bold')

    # Add legend
    ax.legend(wedges, [f'{label}: {count}' for label, count in event_counts.items()],
              title="Distribution", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=8)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_camera_severity_pie(area_data, area_name, area_column):
    """Generate pie chart for Camera vs Severity for a specific area"""
    if 'Severity' not in area_data.columns:
        return ""

    if area_column == 'Camera':
        # Single camera breakdown by severity
        severity_counts = area_data['Severity'].value_counts()
        title = f'Severity Distribution - {area_name}'
    else:
        # Multiple cameras - show severity distribution across area
        severity_counts = area_data['Severity'].value_counts()
        title = f'Severity Distribution - {area_name}'

    if severity_counts.empty:
        return ""

    # Create pie chart
    fig, ax = plt.subplots(figsize=(6, 5))
    # Use minimal professional colors for severity
    colors = PIE_COLORS[:len(severity_counts)]
    if len(severity_counts) > len(PIE_COLORS):
        colors.extend(['#94A3B8'] * (len(severity_counts) - len(PIE_COLORS)))

    wedges, texts, autotexts = ax.pie(
        severity_counts.values,
        labels=None,
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        textprops={'fontsize': 9, 'fontweight': 'bold'}
    )

    ax.set_title(title, fontsize=12, fontweight='bold')

    # Add legend
    ax.legend(wedges, [f'{label}: {count}' for label, count in severity_counts.items()],
              title="Severity Levels", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=8)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=100)
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_weekly_trend(df):
    """Generate weekly trend graph based on entire Excel data"""
    if df.empty or 'Date' not in df.columns:
        return ""

    # Convert Date column to datetime
    df_copy = df.copy()
    df_copy['Date'] = pd.to_datetime(
        df_copy['Date'], dayfirst=True, errors='coerce')
    df_copy = df_copy.dropna(subset=['Date'])

    if df_copy.empty:
        return ""

    # Calculate date range to determine number of weeks
    min_date = df_copy['Date'].min()
    max_date = df_copy['Date'].max()
    date_range = (max_date - min_date).days

    # If data spans less than 7 days, don't generate weekly trend
    if date_range < 7:
        return ""

    # Create week numbers starting from the first week
    df_copy['Week'] = ((df_copy['Date'] - min_date).dt.days // 7) + 1

    # Group by week and count events
    weekly_counts = df_copy.groupby('Week')['Events'].count().reset_index()
    weekly_counts.columns = ['Week', 'Event_Count']

    # Create complete week range
    max_week = weekly_counts['Week'].max()
    complete_weeks = pd.DataFrame({'Week': range(1, max_week + 1)})
    complete_weekly = complete_weeks.merge(
        weekly_counts, on='Week', how='left')
    complete_weekly['Event_Count'] = complete_weekly['Event_Count'].fillna(0)

    # Plot
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(
        complete_weekly['Week'],
        complete_weekly['Event_Count'],
        linewidth=2,
        color='blue',
        alpha=0.7)

    # Add scatter points for better visibility
    ax.scatter(
        complete_weekly['Week'],
        complete_weekly['Event_Count'],
        color='blue',
        s=60,
        zorder=5)

    # Add value labels on points
    for week, count in zip(
            complete_weekly['Week'], complete_weekly['Event_Count']):
        if count > 0:
            ax.annotate(str(int(count)), (week, count), textcoords="offset points",
                        xytext=(0, 10), ha='center', fontsize=9, fontweight='bold')

    ax.set_title(
        'Weekly Trend - Events Per Week',
        fontsize=14,
        fontweight='bold')
    ax.set_xlabel('Week Number')
    ax.set_ylabel('Number of Events')
    ax.set_xticks(complete_weekly['Week'])
    ax.set_xticklabels(
        [f'Week {int(w)}' for w in complete_weekly['Week']], rotation=45)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_monthly_trend(df):
    """Generate monthly trend graph based on entire Excel data"""
    if df.empty or 'Date' not in df.columns:
        return ""

    # Convert Date column to datetime
    df_copy = df.copy()
    df_copy['Date'] = pd.to_datetime(
        df_copy['Date'], dayfirst=True, errors='coerce')
    df_copy = df_copy.dropna(subset=['Date'])

    if df_copy.empty:
        return ""

    # Extract unique months and years
    df_copy['YearMonth'] = df_copy['Date'].dt.to_period('M')
    unique_months = sorted(df_copy['YearMonth'].unique())

    # If data spans only one month, don't generate monthly trend
    if len(unique_months) <= 1:
        return ""

    # Group by month and count events
    monthly_counts = df_copy.groupby(
        'YearMonth')['Events'].count().reset_index()
    monthly_counts.columns = ['YearMonth', 'Event_Count']

    # Create month labels (Month1, Month2, etc.)
    month_labels = [f'Month {i+1}' for i in range(len(unique_months))]
    monthly_counts['Month_Label'] = month_labels

    # Plot
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(
        range(
            len(monthly_counts)),
        monthly_counts['Event_Count'],
        linewidth=2,
        color=PROFESSIONAL_BLUE,
        alpha=0.7)

    # Add scatter points for better visibility
    ax.scatter(
        range(
            len(monthly_counts)),
        monthly_counts['Event_Count'],
        color=PROFESSIONAL_BLUE,
        s=60,
        zorder=5)

    # Add value labels on points
    for i, count in enumerate(monthly_counts['Event_Count']):
        ax.annotate(str(int(count)), (i, count), textcoords="offset points",
                    xytext=(0, 10), ha='center', fontsize=9, fontweight='bold')
    
    ax.set_title(
        'Monthly Trend - Events Per Month',
        fontsize=14,
        fontweight='bold')
    ax.set_xlabel('Month Number')
    ax.set_ylabel('Number of Events')
    ax.set_xticks(range(len(monthly_counts)))
    ax.set_xticklabels(month_labels, rotation=45)

    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64

# Heatmap table functions restored for standalone table display


def create_heatmap_table(df_pivot):
    """Create HTML table with heatmap styling for camera-event data"""
    # Exclude 'Total' column and 'Grand Total' row from heatmap calculation
    df_for_scaling = df_pivot.copy()
    if 'Total' in df_for_scaling.columns:
        df_for_scaling = df_for_scaling.drop('Total', axis=1)
    if 'Grand Total' in df_for_scaling.index:
        df_for_scaling = df_for_scaling.drop('Grand Total', axis=0)

    # Calculate max and min values for color scaling across entire table
    # (excluding totals)
    vmax = df_for_scaling.max().max() if not df_for_scaling.empty else 1
    vmin = df_for_scaling.min().min() if not df_for_scaling.empty else 0

    html = '''
    <div style="overflow-x: auto; margin: 20px 0;">
        <table style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;">
            <thead>
                <tr style="background-color: #f8f9fa;">
                    <th style="border: 1px solid #dee2e6; padding: 8px; text-align: left; font-weight: bold;">Camera</th>
    '''

    # Add column headers
    for col in df_pivot.columns:
        html += f'<th style="border: 1px solid #dee2e6; padding: 8px; text-align: center; font-weight: bold;">{col}</th>'
    html += '</tr></thead><tbody>'

    # Add data rows
    for camera in df_pivot.index:
        html += f'<tr><td style="border: 1px solid #dee2e6; padding: 8px; font-weight: bold;">{camera}</td>'
        for col in df_pivot.columns:
            value = df_pivot.loc[camera, col]

            # Don't apply heatmap to Total columns or Grand Total rows
            if col == 'Total' or camera == 'Grand Total':
                bg_color = '#f8f9fa'
                text_color = '#000'
                html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center; background-color:{bg_color};color:{text_color};font-weight:bold;">{value}</td>'
            else:
                # Calculate color using entire table scaling
                if vmax > vmin and value > 0:
                    intensity = (value - vmin) / (vmax - vmin)
                else:
                    intensity = 0

                # Color scale from white (0) to red (max)
                if value == 0:
                    bg_color = '#ffffff'
                    text_color = '#6c757d'
                else:
                    # Red intensity based on value
                    red = 255
                    green = int(255 - (intensity * 153))  # 255 to 102
                    blue = int(255 - (intensity * 153))   # 255 to 102
                    bg_color = f'rgb({red}, {green}, {blue})'
                    text_color = '#000000' if intensity < 0.7 else '#ffffff'
                html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center; background-color: {bg_color}; color: {text_color}; font-weight: bold;">{value}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'
    return html


def create_event_heatmap_table(df_pivot):
    """Create HTML table with heatmap styling for event-camera data"""
    # Exclude 'Total' column and 'Grand Total' row from heatmap calculation
    df_for_scaling = df_pivot.copy()
    if 'Total' in df_for_scaling.columns:
        df_for_scaling = df_for_scaling.drop('Total', axis=1)
    if 'Grand Total' in df_for_scaling.index:
        df_for_scaling = df_for_scaling.drop('Grand Total', axis=0)

    # Calculate max and min values for color scaling across entire table
    # (excluding totals)
    vmax = df_for_scaling.max().max() if not df_for_scaling.empty else 1
    vmin = df_for_scaling.min().min() if not df_for_scaling.empty else 0

    html = '''
    <div style="overflow-x: auto; margin: 20px 0;">
        <table style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;">
            <thead>
                <tr style="background-color: #f8f9fa;">
                    <th style="border: 1px solid #dee2e6; padding: 8px; text-align: left; font-weight: bold;">Event Type</th>
    '''

    # Add column headers (camera names)
    for col in df_pivot.columns:
        html += f'<th style="border: 1px solid #dee2e6; padding: 8px; text-align: center; font-weight: bold;">{col}</th>'
    html += '</tr></thead><tbody>'

    # Add data rows (event types)
    for event in df_pivot.index:
        html += f'<tr><td style="border: 1px solid #dee2e6; padding: 8px; font-weight: bold;">{event}</td>'
        for col in df_pivot.columns:
            value = df_pivot.loc[event, col]

            # Don't apply heatmap to Total columns or Grand Total rows
            if col == 'Total' or event == 'Grand Total':
                html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center; background-color:#f8f9fa;color:#000;font-weight:bold;">{value}</td>'
            else:
                # Calculate color using entire table scaling
                if vmax > vmin and value > 0:
                    intensity = (value - vmin) / (vmax - vmin)
                else:
                    intensity = 0

                # Color scale from white (0) to red (max)
                if value == 0:
                    bg_color = '#ffffff'
                    text_color = '#6c757d'
                else:
                    # Red intensity based on value
                    red = 255
                    green = int(255 - (intensity * 153))  # 255 to 102
                    blue = int(255 - (intensity * 153))   # 255 to 102
                    bg_color = f'rgb({red}, {green}, {blue})'
                    text_color = '#000000' if intensity < 0.7 else '#ffffff'

                html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center; background-color: {bg_color}; color: {text_color}; font-weight: bold;">{value}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'
    return html


def create_high_risk_heatmap_table(high_risk_table):
    """Create HTML table with heatmap styling for high risk events data"""
    if high_risk_table.empty:
        return ""
    
    # Calculate max and min values for color scaling
    vmax = high_risk_table['Count'].max() if not high_risk_table.empty else 1
    vmin = high_risk_table['Count'].min() if not high_risk_table.empty else 0
    
    html = '''
    <div style="overflow-x: auto; margin: 20px 0;">
        <table style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;">
            <thead>
                <tr style="background-color: #f8f9fa;">
                    <th style="border: 1px solid #dee2e6; padding: 8px; text-align: left; font-weight: bold;">Severity</th>
                    <th style="border: 1px solid #dee2e6; padding: 8px; text-align: left; font-weight: bold;">Event Type</th>
                    <th style="border: 1px solid #dee2e6; padding: 8px; text-align: center; font-weight: bold;">Count</th>
                    <th style="border: 1px solid #dee2e6; padding: 8px; text-align: center; font-weight: bold;">% of Total</th>
                </tr>
            </thead>
            <tbody>
    '''
    
    # Add data rows with heatmap styling
    for _, row in high_risk_table.iterrows():
        severity = row['Severity']
        event_type = row['Events']
        count = row['Count']
        percentage = row['Percentage']
        
        # Severity color coding
        severity_colors = {
            'Critical': '#dc2626',
            'High': '#ea580c',
            'Severe': '#ca8a04',
            'Major': '#ca8a04'
        }
        severity_color = severity_colors.get(severity, '#6b7280')
        
        # Calculate color intensity for Count column
        if vmax > vmin and count > 0:
            intensity = (count - vmin) / (vmax - vmin)
        else:
            intensity = 0
        
        # Color scale from white (0) to red (max)
        if count == 0:
            bg_color = '#ffffff'
            text_color = '#6c757d'
        else:
            # Red intensity based on value
            red = 255
            green = int(255 - (intensity * 153))  # 255 to 102
            blue = int(255 - (intensity * 153))   # 255 to 102
            bg_color = f'rgb({red}, {green}, {blue})'
            text_color = '#000000' if intensity < 0.7 else '#ffffff'
        
        html += f'<tr><td style="border: 1px solid #dee2e6; padding: 8px; color: {severity_color}; font-weight: bold;">{severity}</td>'
        html += f'<td style="border: 1px solid #dee2e6; padding: 8px; color: #374151;">{event_type}</td>'
        html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center; background-color: {bg_color}; color: {text_color}; font-weight: bold;">{count}</td>'
        html += f'<td style="border: 1px solid #dee2e6; padding: 8px; text-align: center; color: #6b7280;">{percentage}</td></tr>'
    
    html += '</tbody></table></div>'
    return html


def get_day_name(day_number):
    """
    Convert day number to day name (Sunday=0, Monday=1, ..., Saturday=6)
    """
    day_names = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    return day_names[day_number]

def get_smart_date_display(date_range, max_labels=30, is_weekly=False):
    """
    Smart date display logic for x-axis labels
    
    Args:
        date_range: List of dates or date labels
        max_labels: Maximum number of labels to show (default 30)
        is_weekly: If True, convert dates to day names for weekly reports
    
    Returns:
        tuple: (positions_to_show, labels_to_show)
    """
    if is_weekly and len(date_range) <= 7:
        # For weekly reports with 7 days or fewer, show day names
        day_names = []
        for date_str in date_range:
            try:
                # Try to parse the date and get day name
                date_obj = pd.to_datetime(date_str, dayfirst=True)
                day_number = (date_obj.dayofweek + 1) % 7  # Convert to Sunday=0
                day_names.append(get_day_name(day_number))
            except:
                day_names.append(date_str)
        return list(range(len(day_names))), day_names
    
    total_days = len(date_range)
    
    if total_days <= 31:
        # Show all days if 31 or fewer
        return list(range(total_days)), date_range
    
    elif total_days < 60:
        # For 32-59 days: skip every other day to show ~30 days
        step = 2
        positions = list(range(0, total_days, step))
        labels = [date_range[i] for i in positions]
        return positions, labels
    
    elif total_days >= 60:
        # For 60+ days: use smart stepping to show exactly max_labels
        if total_days <= 90:  # 2-3 months
            step = max(2, total_days // max_labels)
        elif total_days <= 180:  # 3-6 months
            step = max(7, total_days // max_labels)
        else:  # 6+ months
            step = max(14, total_days // max_labels)
        
        positions = list(range(0, total_days, step))
        # Ensure we don't exceed max_labels
        if len(positions) > max_labels:
            positions = positions[:max_labels]
        
        labels = [date_range[i] for i in positions]
        return positions, labels
    
    return list(range(total_days)), date_range


def plot_events_per_day_range(df, start_date, end_date, date_range_text, is_weekly=False):
    # Validate data before processing
    if df.empty:
        return ""
    
    # Group by date and count events
    events_per_day = df.groupby(df['Date'].dt.date)['Events'].count().reset_index()
    events_per_day.columns = ['Date', 'Event_Count']
    
    # Create complete date range
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    complete_dates = pd.DataFrame({'Date': date_range.date})
    complete_events = complete_dates.merge(events_per_day, on='Date', how='left')
    complete_events['Event_Count'] = complete_events['Event_Count'].fillna(0)
    
    # Plot
    fig, ax = plt.subplots(figsize=(12, 5))
    
    if len(complete_events) > 0:
        complete_events = complete_events.sort_values('Date')
        
        # Plot line connecting all points (including zeros)
        ax.plot(range(len(complete_events)), complete_events['Event_Count'],
                linewidth=2, color='purple', alpha=0.7)
        
        # Add scatter points for non-zero events only
        non_zero_mask = complete_events['Event_Count'] > 0
        non_zero_positions = [i for i, is_nonzero in enumerate(non_zero_mask) if is_nonzero]
        non_zero_counts = complete_events[complete_events['Event_Count'] > 0]['Event_Count']
        
        if len(non_zero_positions) > 0:
            ax.scatter(non_zero_positions, non_zero_counts, 
                      color='purple', s=60, zorder=5, label='With Events')
            
            # Add value labels for non-zero points only
            for pos, count in zip(non_zero_positions, non_zero_counts):
                ax.annotate(str(int(count)), (pos, count), 
                           textcoords="offset points", xytext=(0, 10), 
                           ha='center', fontsize=8, fontweight='bold')
        
        # Use smart date display logic with day names for weekly reports
        if is_weekly and len(complete_events) <= 7:
            # For weekly reports, use day names
            day_labels = []
            for date in complete_events['Date']:
                day_number = (pd.to_datetime(date).dayofweek + 1) % 7  # Convert to Sunday=0
                day_labels.append(get_day_name(day_number))
            positions_to_show = list(range(len(day_labels)))
            labels_to_show = day_labels
        else:
            # For other reports, use date format
            date_labels = [d.strftime('%d-%m-%Y') for d in complete_events['Date']]
            positions_to_show, labels_to_show = get_smart_date_display(date_labels, max_labels=30)
        
        ax.set_xticks(positions_to_show)
        ax.set_xticklabels(labels_to_show, rotation=45)
        ax.legend()
    
    ax.set_title(f'Events Per Day - {date_range_text}', fontsize=14, fontweight='bold')
    ax.set_xlabel('Day' if is_weekly else 'Date')
    ax.set_ylabel('Number of Events')
    
    # Set y-axis to start slightly above 0
    ax.set_ylim(bottom=-2)
    
    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def plot_avg_interval_range(df, start_date, end_date, date_range_text, is_weekly=False):
    # Convert Time column if it exists and has valid data
    if 'Time' in df.columns and not df['Time'].isna().all():
        df['Time'] = pd.to_datetime(df['Time'], dayfirst=True, errors='coerce')
        # Remove rows with invalid time data for interval calculation
        df_time_filtered = df.dropna(subset=['Time'])
    else:
        df_time_filtered = pd.DataFrame()  # Empty dataframe if no valid time data

    daily_intervals = []
    for date in df['Date'].dt.date.unique():
        if not df_time_filtered.empty:
            day_data = df_time_filtered[df_time_filtered['Date'].dt.date == date].sort_values('Time')
            if len(day_data) > 1:
                time_diffs = [
                    (day_data.iloc[i]['Time'] - day_data.iloc[i - 1]['Time']).total_seconds() / 60
                    for i in range(1, len(day_data))
                ]
                avg_interval = sum(time_diffs) / len(time_diffs)
                daily_intervals.append({'Date': date, 'Avg_Interval_Minutes': avg_interval})
            else:
                daily_intervals.append({'Date': date, 'Avg_Interval_Minutes': 0})
        else:
            daily_intervals.append({'Date': date, 'Avg_Interval_Minutes': 0})

    intervals_df = pd.DataFrame(daily_intervals)

    # Create complete date range
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    complete_dates = pd.DataFrame({'Date': date_range.date})
    complete_intervals = complete_dates.merge(intervals_df, on='Date', how='left')
    complete_intervals['Avg_Interval_Minutes'] = complete_intervals['Avg_Interval_Minutes'].fillna(0)

    fig, ax = plt.subplots(figsize=(12, 5))

    # Plot all intervals (including zeros)
    if len(complete_intervals) > 0:
        complete_intervals = complete_intervals.sort_values('Date')
        
        # Plot line connecting all points (including zeros)
        ax.plot(range(len(complete_intervals)), complete_intervals['Avg_Interval_Minutes'],
                linewidth=2, color="teal", alpha=0.7)
        
        # Add scatter points for non-zero intervals only
        non_zero_mask = complete_intervals['Avg_Interval_Minutes'] > 0
        non_zero_positions = [i for i, is_nonzero in enumerate(non_zero_mask) if is_nonzero]
        non_zero_values = complete_intervals[complete_intervals['Avg_Interval_Minutes'] > 0]['Avg_Interval_Minutes']
        
        if len(non_zero_positions) > 0:
            ax.scatter(non_zero_positions, non_zero_values, 
                      color="teal", s=60, zorder=5, label='Multiple Events')
            
            # Add value labels for non-zero points only
            for pos, val in zip(non_zero_positions, non_zero_values):
                ax.annotate(f'{val:.1f}m', (pos, val), 
                           textcoords="offset points", xytext=(0, 10), 
                           ha='center', fontsize=8, fontweight='bold')
        
        # Use smart date display logic with day names for weekly reports
        if is_weekly and len(complete_intervals) <= 7:
            # For weekly reports, use day names
            day_labels = []
            for date in complete_intervals['Date']:
                day_number = (pd.to_datetime(date).dayofweek + 1) % 7  # Convert to Sunday=0
                day_labels.append(get_day_name(day_number))
            positions_to_show = list(range(len(day_labels)))
            labels_to_show = day_labels
        else:
            # For other reports, use date format
            date_labels = [d.strftime('%d-%m-%Y') for d in complete_intervals['Date']]
            positions_to_show, labels_to_show = get_smart_date_display(date_labels, max_labels=30)
        
        ax.set_xticks(positions_to_show)
        ax.set_xticklabels(labels_to_show, rotation=45)
        ax.legend()

    ax.set_title(f'Average Time Interval Between Events - {date_range_text}', fontsize=14, fontweight='bold')
    ax.set_xlabel('Day' if is_weekly else 'Date')
    ax.set_ylabel('Average Interval (Minutes)')
    
    # Set y-axis to start slightly above 0
    ax.set_ylim(bottom=-2)
    
    plt.tight_layout()
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    return img_base64


def display_latest_videos_section(df_selected, excel_file=None):
    """
    Display the Top Videos section prioritized by severity and recency
    Now supports extracting and displaying actual hyperlinks from Excel files
    """
    try:
        st.markdown("### 📹 Top Videos by Priority")
        st.markdown("*Highest priority video recordings ranked by severity (Critical > High > Moderate > Low) and recency*")
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Extract top video links from the selected date range data with hyperlink support
        top_videos = extract_video_links(df_selected, top_n=8, excel_file=excel_file)
        
        if not top_videos:
            st.info("📌 No video paths found in the current dataset or date range.")
            return
    except Exception as e:
        st.error(f"❌ Error processing video data: {e}")
        return
    
    try:
        # Display top videos in priority order
        cols = st.columns(3)
        
        # Define severity colors and icons
        severity_styles = {
            'Critical': {'color': '#dc2626', 'bg': '#fee2e2', 'icon': '🚨'},
            'High': {'color': '#ea580c', 'bg': '#fed7aa', 'icon': '⚠️'},
            'Moderate': {'color': '#ca8a04', 'bg': '#fef3c7', 'icon': '⚡'},
            'Medium': {'color': '#ca8a04', 'bg': '#fef3c7', 'icon': '⚡'},
            'Low': {'color': '#16a34a', 'bg': '#dcfce7', 'icon': '📝'},
            'Unknown': {'color': '#6b7280', 'bg': '#f3f4f6', 'icon': '❓'}
        }
        
        for i, (event_type, video_info) in enumerate(top_videos.items()):
            col_index = i % len(cols)
            
            with cols[col_index]:
                try:
                    # Create priority rank
                    rank = i + 1
                    rank_colors = ['#ffd700', '#c0c0c0', '#cd7f32', '#4a90e2', '#9b59b6', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#34495e', '#e67e22', '#3498db', '#8e44ad', '#16a085', '#f1c40f', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6']  # Extended colors for more videos
                    rank_color = rank_colors[min(i, len(rank_colors)-1)]
                    
                    # Safely format the date
                    try:
                        date_str = video_info['date'].strftime('%d-%m-%Y %H:%M')
                    except:
                        date_str = str(video_info['date'])
                    
                    # Safely get video info with defaults
                    camera = str(video_info.get('camera', 'Unknown'))
                    area = str(video_info.get('area', 'Unknown'))
                    severity = str(video_info.get('severity', 'Unknown'))
                    
                    # Get severity styling
                    severity_style = severity_styles.get(severity, severity_styles['Unknown'])
                    
                    # Create video container with header
                    st.markdown(f"""
                    <div style="background-color: white; border: 2px solid {rank_color}; border-radius: 12px; padding: 15px; margin: 10px 0; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px;">
                            <div style="display: flex; align-items: center;">
                                <div style="background-color: {rank_color}; color: white; border-radius: 50%; width: 25px; height: 25px; display: flex; align-items: center; justify-content: center; margin-right: 8px; font-size: 12px; font-weight: bold;">
                                    #{rank}
                                </div>
                                <h4 style="margin: 0; color: #1f2937; font-size: 16px; font-weight: bold;">{event_type}</h4>
                            </div>
                            <div style="background-color: {severity_style['bg']}; color: {severity_style['color']}; padding: 3px 6px; border-radius: 8px; font-size: 11px; font-weight: bold;">
                                {severity_style['icon']} {severity}
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Get video path for player
                    video_path = video_info.get('video_path', '')
                    
                    if isinstance(video_path, str) and video_path:
                        # Try to embed video player
                        video_extensions = ['.mp4', '.webm', '.ogg', '.mov', '.avi', '.mkv', '.flv', '.wmv', '.m4v']
                        
                        # Check if it's a video file or video URL (including Azure Blob Storage)
                        is_video_file = any(video_path.lower().endswith(ext) for ext in video_extensions)
                        is_video_url = (video_path.startswith('http') or video_path.startswith('https')) and (
                            is_video_file or 
                            any(ext in video_path.lower() for ext in video_extensions) or
                            'blob.core.windows.net' in video_path.lower()  # Azure Blob Storage
                        )
                        
                        if is_video_file or is_video_url:
                            try:
                                # Try to display as embedded video player
                                st.video(video_path)
                                
                                # Always provide option to open in new tab
                                st.markdown(f"""
                                <div style="text-align: center; margin: 10px 0;">
                                    <a href="{video_path}" target="_blank" rel="noopener noreferrer" style="background-color: #10b981; color: white; padding: 8px 16px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: bold; display: inline-block;">
                                        🚀 Open in New Tab
                                    </a>
                                </div>
                                """, unsafe_allow_html=True)
                            except Exception as e:
                                # Show error but still provide clickable link
                                st.markdown(f"""
                                <div style="background-color: #fef3c7; border: 1px solid #f59e0b; border-radius: 8px; padding: 15px; margin: 10px 0; text-align: center;">
                                    <div style="color: #92400e; font-size: 14px; margin-bottom: 10px;">⚠️ Video player could not load, but you can access the link directly:</div>
                                    <a href="{video_path}" target="_blank" rel="noopener noreferrer" style="background-color: #3b82f6; color: white; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold; display: inline-block; margin: 5px;">
                                        🎥 Open Video in New Tab
                                    </a>
                                </div>
                                """, unsafe_allow_html=True)
                        else:
                            # Not detected as video but might still be a video link - provide clickable link
                            if video_path.startswith('http') or video_path.startswith('https'):
                                st.markdown(f"""
                                <div style="background-color: #f0f9ff; border: 1px solid #0ea5e9; border-radius: 8px; padding: 15px; margin: 10px 0; text-align: center;">
                                    <div style="color: #0369a1; font-size: 14px; margin-bottom: 10px;">🌐 Click to open link in new tab:</div>
                                    <a href="{video_path}" target="_blank" rel="noopener noreferrer" style="background-color: #0ea5e9; color: white; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-size: 14px; font-weight: bold; display: inline-block; margin: 5px;">
                                        🎥 Open Video in New Tab
                                    </a>
                                </div>
                                """, unsafe_allow_html=True)
                            else:
                                st.markdown(f"""
                                <div style="background-color: #f3f4f6; border: 1px solid #d1d5db; border-radius: 8px; padding: 15px; margin: 10px 0; text-align: center;">
                                    <div style="color: #6b7280; font-size: 14px; margin-bottom: 10px;">📁 Local file path</div>
                                    <div style="background-color: #e5e7eb; padding: 8px 12px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; word-break: break-all;">
                                        {video_path}
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)
                    else:
                        # No video path available
                        st.markdown(f"""
                        <div style="background-color: #fef2f2; border: 2px dashed #fca5a5; border-radius: 8px; padding: 40px; text-align: center; margin: 10px 0;">
                            <div style="font-size: 48px; margin-bottom: 10px;">❌</div>
                            <div style="color: #dc2626; font-size: 14px;">No Video Hyperlink Available</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Add metadata below video
                    st.markdown(f"""
                    <div style="background-color: #f8fafc; padding: 10px; border-radius: 6px; margin-top: 10px;">
                        <div style="color: #6b7280; font-size: 12px;">
                            <strong>📅</strong> {date_str} &nbsp;&nbsp;
                            <strong>📷</strong> {camera} &nbsp;&nbsp;
                            <strong>📍</strong> {area}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                except Exception as e:
                    st.error(f"❌ Error displaying video for {event_type}: {e}")
        
        # Simple summary information
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Count videos by severity for summary
        severity_counts = {}
        hyperlink_count = 0
        for video_info in top_videos.values():
            severity = video_info.get('severity', 'Unknown')
            severity_counts[severity] = severity_counts.get(severity, 0) + 1
            if video_info.get('has_hyperlink', False):
                hyperlink_count += 1
        
        # Create severity breakdown
        severity_breakdown = []
        for severity, count in severity_counts.items():
            severity_style = severity_styles.get(severity, severity_styles['Unknown'])
            severity_breakdown.append(f"<span style='color: {severity_style['color']};'>{severity_style['icon']} {count} {severity}</span>")
        
        # Simple summary
        hyperlink_status = f" (🔗 {hyperlink_count} with extracted hyperlinks)" if hyperlink_count > 0 else ""
        st.markdown(f"""
        <div style="background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 12px; margin: 15px 0;">
            <p style="color: #0369a1; margin: 0; font-size: 13px;">
                <strong>Top 8 priority videos</strong> by severity and recency{hyperlink_status}
            </p>
        </div>
        """, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"❌ Error displaying video section: {e}")


def main():
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        # Check for required columns
        required_columns = ['Camera', 'Events', 'Area', 'Severity']
        missing_columns = [
            col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"Missing required columns: {', '.join(missing_columns)}")
            return
        
        # Check for optional Event Video Path column
        has_video_paths = 'Event Video Path' in df.columns
        if has_video_paths:
            st.info("📹 Event Video Path column found - Video analysis will be available!")
        else:
            st.info("📹 Event Video Path column not found - Video analysis will be skipped.")

        # Convert Date column and validate
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
        invalid_dates = df['Date'].isna().sum()
        if invalid_dates > 0:
            st.warning(
                f"Found {invalid_dates} rows with invalid dates. These will be excluded from analysis.")

        df = df.dropna(subset=['Date'])
        if df.empty:
            st.error("No valid dates found in the uploaded file.")
            return

        # Validate 'Time' column for interval analysis (optional but
        # recommended)
        if 'Time' not in df.columns:
            st.warning(
                "'Time' column not found. Interval analysis will show 0 for all days.")
            df['Time'] = pd.NaT  # Add empty Time column

        st.success("File uploaded and processed successfully!")

        # Add Custom Date Range Selector
        st.markdown("### 📅 Custom Date Range Filter")
        st.markdown(
            "*Select a custom date range to filter the analysis, or leave default to show all data*")

        # Get date range from data
        min_date = df['Date'].min().date()
        max_date = df['Date'].max().date()

        # Create two columns for start and end date
        col1, col2 = st.columns(2)

        with col1:
            start_date = st.date_input(
                "Start Date",
                value=min_date,
                min_value=min_date,
                max_value=max_date,
                key="start_date"
            )

        with col2:
            end_date = st.date_input(
                "End Date",
                value=max_date,
                min_value=min_date,
                max_value=max_date,
                key="end_date"
            )

        # Validate date range
        if start_date > end_date:
            st.error(
                "Start date cannot be after end date. Please select a valid date range.")
            return

        # Filter dataframe based on selected date range
        df_selected = df[
            (df['Date'].dt.date >= start_date) &
            (df['Date'].dt.date <= end_date)
        ]

        if df_selected.empty:
            st.warning(
                f"No data found for the selected date range ({start_date} to {end_date}).")
            return

        # Display selected date range info
        total_days = (end_date - start_date).days + 1
        date_range_text = f"{start_date.strftime('%B %d, %Y')} to {end_date.strftime('%B %d, %Y')}"
        if start_date == min_date and end_date == max_date:
            date_range_text = "All Available Data"

        st.info(
            f"📊 Analyzing data for: **{date_range_text}** ({total_days} days)")
        st.markdown("<br>", unsafe_allow_html=True)

        # Data Overview Section for selected date range
        st.markdown("## 📊 Data Overview")
        st.markdown(f"*Overview for {date_range_text}*")

        # Calculate overview metrics for selected period
        total_events_overview = len(df_selected)

        # Count unique event types
        if 'Events' in df_selected.columns:
            unique_event_types = df_selected['Events'].nunique()
        else:
            unique_event_types = 0

        # Count areas monitored (prefer Area column, fallback to Camera)
        if 'Area' in df_selected.columns:
            areas_monitored = df_selected['Area'].nunique()
        elif 'Camera' in df_selected.columns:
            areas_monitored = df_selected['Camera'].nunique()
        else:
            areas_monitored = 0

        # Calculate Review Percent and Accuracy Percent for dashboard
        review_percent = 85  # Default value
        if 'Reviewed' in df_selected.columns:
            # Count "Yes" values in the Reviewed column
            reviewed_yes_count = df_selected['Reviewed'].str.contains('yes', case=False, na=False).sum()
            review_percent = round((reviewed_yes_count / len(df_selected)) * 100) if len(df_selected) > 0 else 0
        elif 'Review' in df_selected.columns:
            reviewed_events = df_selected['Review'].notna().sum()
            review_percent = round((reviewed_events / len(df_selected)) * 100) if len(df_selected) > 0 else 0
        elif 'Status' in df_selected.columns:
            reviewed_events = df_selected['Status'].str.contains('review|complete|resolved', case=False, na=False).sum()
            review_percent = round((reviewed_events / len(df_selected)) * 100) if len(df_selected) > 0 else 0

        # Display overview cards (4 columns now)
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"""
            <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <div style="color: #6b7280; font-size: 14px; margin-bottom: 8px;">📈 Total Events</div>
                <div style="color: #111827; font-size: 32px; font-weight: bold;">{total_events_overview:,}</div>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <div style="color: #6b7280; font-size: 14px; margin-bottom: 8px;">⚠️ Event Types</div>
                <div style="color: #111827; font-size: 32px; font-weight: bold;">{unique_event_types}</div>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <div style="color: #6b7280; font-size: 14px; margin-bottom: 8px;">⭕ Areas Monitored</div>
                <div style="color: #111827; font-size: 32px; font-weight: bold;">{areas_monitored}</div>
            </div>
            """, unsafe_allow_html=True)

        with col4:
            st.markdown(f"""
            <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <div style="color: #6b7280; font-size: 14px; margin-bottom: 8px;">✅ Review Percent</div>
                <div style="color: #111827; font-size: 32px; font-weight: bold;">{review_percent}%</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # High Impact Analysis and Top 3 Cameras sections
        st.markdown("<br>", unsafe_allow_html=True)

        # Calculate metrics for the selected month/year
        total_events_selected = len(df_selected)

        # Calculate different event types for High Impact Analysis
        if 'Events' in df_selected.columns:
            # More flexible matching for Near Miss events (handles Near-Miss,
            # NearMiss, Near Miss, etc.)
            near_miss_pattern = r'near[-\s]*miss'
            near_miss_events = df_selected[df_selected['Events'].str.contains(
                near_miss_pattern, case=False, na=False, regex=True)]
            total_near_miss = len(near_miss_events)
            near_miss_percentage = (
                total_near_miss /
                total_events_selected *
                100) if total_events_selected > 0 else 0

            # More flexible matching for Emergency events
            emergency_pattern = r'emergency'
            emergency_events = df_selected[df_selected['Events'].str.contains(
                emergency_pattern, case=False, na=False, regex=True)]
            total_emergency = len(emergency_events)
            emergency_percentage = (
                total_emergency /
                total_events_selected *
                100) if total_events_selected > 0 else 0

            # High Risk Events (more flexible severity matching)
            if 'Severity' in df_selected.columns:
                # Convert severity to string and make case-insensitive
                # comparison
                severity_str = df_selected['Severity'].astype(str).str.lower()
                high_risk_pattern = r'(critical|high|severe|major)'
                high_risk_mask = severity_str.str.contains(
                    high_risk_pattern, case=False, na=False, regex=True)
                total_high_risk = high_risk_mask.sum()
            else:
                total_high_risk = 0
        else:
            total_near_miss = 0
            near_miss_percentage = 0
            total_emergency = 0
            emergency_percentage = 0
            total_high_risk = 0

        # Calculate top cameras for selected period
        camera_column = None
        if 'Camera' in df_selected.columns:
            camera_column = 'Camera'
        elif 'Area' in df_selected.columns:
            camera_column = 'Area'

        if camera_column:
            camera_counts = df_selected[camera_column].value_counts().head(3)
            camera_contributions = []
            for camera, count in camera_counts.items():
                # Calculate high severity events for this camera with flexible
                # matching
                if 'Severity' in df_selected.columns:
                    camera_data = df_selected[df_selected[camera_column] == camera]
                    severity_str = camera_data['Severity'].astype(
                        str).str.lower()
                    high_severity_pattern = r'(critical|high|severe|major)'
                    camera_high_severity = severity_str.str.contains(
                        high_severity_pattern, case=False, na=False, regex=True).sum()
                else:
                    camera_high_severity = 0

                # Extract location/plant info from camera name if available
                location = "Unknown Location"
                if isinstance(camera, str):
                    if 'plant' in camera.lower():
                        # Extract plant info (e.g., "Plant-2 Bagging Station"
                        # -> "Plant-2 Bagging")
                        plant_match = re.search(
                            r'plant[-\s]*\d+[-\s]*\w*', camera, re.IGNORECASE)
                        if plant_match:
                            location = plant_match.group(0)
                        else:
                            location = "Plant Location"
                    elif any(keyword in camera.lower() for keyword in ['area', 'zone', 'section', 'dept']):
                        location = "Operational Area"
                    else:
                        location = "Monitoring Point"

                camera_contributions.append({
                    'camera': camera,
                    'total': count,
                    'high_severity': camera_high_severity,
                    'location': location
                })
        else:
            camera_contributions = []

        # High Impact Analysis Section
        st.markdown("### High Impact Analysis")
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(f"""
            <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: left;">
                <h4 style="color: #374151; margin-top: 0; margin-bottom: 10px; font-size: 18px;">Near-Miss Events</h4>
                <div style="color: #dc2626; font-size: 24px; font-weight: bold; margin-bottom: 5px;">{total_near_miss}</div>
                <div style="color: #6b7280; font-size: 12px;">{near_miss_percentage:.1f}% of total events</div>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: left;">
                <h4 style="color: #374151; margin-top: 0; margin-bottom: 10px; font-size: 18px;">Emergency Events</h4>
                <div style="color: #dc2626; font-size: 24px; font-weight: bold; margin-bottom: 5px;">{total_emergency}</div>
                <div style="color: #6b7280; font-size: 12px;">{emergency_percentage:.1f}% of total events</div>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: left;">
                <h4 style="color: #374151; margin-top: 0; margin-bottom: 10px; font-size: 18px;">High Risk Events</h4>
                <div style="color: #dc2626; font-size: 24px; font-weight: bold; margin-bottom: 5px;">{total_high_risk}</div>
                <div style="color: #6b7280; font-size: 12px;">Combined critical incidents</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # High Risk Events Table Section
        st.markdown("### 🚨 High Risk Events Analysis")
        st.markdown("*Detailed breakdown of high-risk events by severity and type*")
        
        if 'Severity' in df_selected.columns and 'Events' in df_selected.columns:
            # Filter high risk events
            severity_str = df_selected['Severity'].astype(str).str.lower()
            high_risk_pattern = r'(critical|high|severe|major)'
            high_risk_mask = severity_str.str.contains(high_risk_pattern, case=False, na=False, regex=True)
            high_risk_events = df_selected[high_risk_mask]
            
            if not high_risk_events.empty:
                # Create high risk events table
                high_risk_table = high_risk_events.groupby(['Severity', 'Events']).size().reset_index(name='Count')
                high_risk_table = high_risk_table.sort_values(['Severity', 'Count'], ascending=[True, False])
                
                # Calculate percentage for each row
                total_events = len(df_selected)
                high_risk_table['Percentage'] = (high_risk_table['Count'] / total_events * 100).round(1)
                high_risk_table['Percentage'] = high_risk_table['Percentage'].astype(str) + '%'
                
                # Create heatmap-styled table for High Risk Events using the dedicated function
                heatmap_html = create_high_risk_heatmap_table(high_risk_table)
                st.markdown(heatmap_html, unsafe_allow_html=True)
                
                # Show summary
                high_risk_percentage = (total_high_risk / total_events * 100) if total_events > 0 else 0
                st.markdown(f"""
                <div style="margin-top: 15px; padding: 10px; background-color: #fef2f2; border-radius: 6px; border-left: 4px solid #dc2626;">
                    <p style="margin: 0; color: #dc2626; font-size: 14px; font-weight: bold;">
                        ⚠️ Total High Risk Events: {total_high_risk} ({high_risk_percentage:.1f}% of all events)
                    </p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style="background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 20px; text-align: center;">
                    <p style="color: #0369a1; margin: 0;">✅ No high-risk events found in the selected period</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="background-color: #fef3c7; border: 1px solid #f59e0b; border-radius: 8px; padding: 20px; text-align: center;">
                <p style="color: #92400e; margin: 0;">⚠️ Severity column not available - cannot analyze high-risk events</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Top 3 Contributing Cameras Section
        st.markdown("### Top 3 Contributing Cameras")

        if camera_contributions:
            col1, col2, col3 = st.columns(3)
            columns = [col1, col2, col3]

            for i, (col, cam) in enumerate(
                    zip(columns, camera_contributions), 1):
                with col:
                    st.markdown(f"""
                    <div style="background-color: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px;">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                            <span style="background-color: #7c3aed; color: white; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;"># {i}</span>
                            <span style="background-color: #f3f4f6; color: #6b7280; padding: 4px 8px; border-radius: 12px; font-size: 12px;">{cam['high_severity']} High Severity</span>
                        </div>
                        <h4 style="color: #7c3aed; margin: 10px 0 5px 0; font-size: 16px; font-weight: bold;">{cam['camera']}</h4>
                        <div style="color: #6b7280; font-size: 12px; margin-bottom: 10px;">{cam['location']}</div>
                        <div style="color: #7c3aed; font-size: 20px; font-weight: bold;">{cam['total']} events</div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; text-align: center;">
                <p style="color: #6b7280; margin: 0;">No camera data available for this period</p>
            </div>
            """, unsafe_allow_html=True)

        # Top 10 Cameras by Event Count Section
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### Top 10 Cameras by Event Count")

        if camera_column and len(df_selected) > 0:
            # Create pivot table for camera event breakdown
            if 'Events' in df_selected.columns:
                # Get top 10 cameras by total event count
                top_cameras = df_selected[camera_column].value_counts().head(
                    10)
                top_camera_names = top_cameras.index.tolist()

                # Get top 5 event types by total count from all data (not
                # filtered)
                top_events_all = df_selected['Events'].value_counts().head(5)
                top_event_names_all = top_events_all.index.tolist()

                # Create pivot table from ALL data first (not filtering by
                # cameras yet)
                full_camera_pivot = df_selected.groupby(
                    [camera_column, 'Events']).size().unstack(fill_value=0)

                if not full_camera_pivot.empty:
                    # Filter to only show top 10 cameras as rows
                    camera_event_pivot = full_camera_pivot.loc[full_camera_pivot.index.intersection(
                        top_camera_names)]

                    # Ensure we have all top 10 cameras as rows (add missing
                    # ones with zeros)
                    for camera in top_camera_names:
                        if camera not in camera_event_pivot.index:
                            # Add row of zeros for this camera
                            camera_event_pivot.loc[camera] = 0

                    # Reorder rows by total count (descending)
                    camera_event_pivot['Total'] = camera_event_pivot.sum(
                        axis=1)
                    camera_event_pivot = camera_event_pivot.sort_values(
                        'Total', ascending=False).drop('Total', axis=1)

                    # Ensure we have all top 5 event types as columns (add
                    # missing ones with zeros)
                    for event in top_event_names_all:
                        if event not in camera_event_pivot.columns:
                            camera_event_pivot[event] = 0

                    # Keep only top 5 event types and reorder by their total
                    # counts
                    available_top_events = [
                        evt for evt in top_event_names_all if evt in camera_event_pivot.columns]
                    camera_event_pivot = camera_event_pivot[available_top_events]

                    # Display the heatmap table
                    if not camera_event_pivot.empty:
                        heatmap_html = create_heatmap_table(camera_event_pivot)
                        st.markdown(heatmap_html, unsafe_allow_html=True)
                    else:
                        st.info(
                            "No event data available for cameras in the selected period.")
                else:
                    st.info(
                        "No camera event data available for the selected period.")
            else:
                st.info("Events column not found in the data.")
        else:
            st.info("No camera/area data available for the selected period.")

        # Top 5 Event Types Section
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### Top 5 Event Types")

        if camera_column and len(df_selected) > 0:
            if 'Events' in df_selected.columns:
                # Get top 5 event types by count
                top_events = df_selected['Events'].value_counts().head(5)
                top_event_names = top_events.index.tolist()

                # Get top 5 cameras by total event count from all data (not
                # filtered)
                top_cameras_all = df_selected[camera_column].value_counts().head(
                    5)
                top_camera_names_all = top_cameras_all.index.tolist()

                # Create pivot table from ALL data first (not filtering by
                # event types yet)
                full_pivot = df_selected.groupby(
                    ['Events', camera_column]).size().unstack(fill_value=0)

                if not full_pivot.empty:
                    # Filter to only show top 5 event types as rows
                    event_camera_pivot = full_pivot.loc[full_pivot.index.intersection(
                        top_event_names)]

                    # Ensure we have all top 5 events as rows (add missing ones
                    # with zeros)
                    for event in top_event_names:
                        if event not in event_camera_pivot.index:
                            # Add row of zeros for this event
                            event_camera_pivot.loc[event] = 0

                    # Reorder rows by event frequency (as per top_event_names
                    # order)
                    event_camera_pivot = event_camera_pivot.reindex(
                        top_event_names)

                    # Ensure we have all top 5 cameras as columns (add missing
                    # ones with zeros)
                    for camera in top_camera_names_all:
                        if camera not in event_camera_pivot.columns:
                            event_camera_pivot[camera] = 0

                    # Keep only top 5 cameras and reorder by their total counts
                    available_top_cameras = [
                        cam for cam in top_camera_names_all if cam in event_camera_pivot.columns]
                    event_camera_pivot = event_camera_pivot[available_top_cameras]

                    # Display the event types heatmap table
                    if not event_camera_pivot.empty:
                        event_heatmap_html = create_event_heatmap_table(
                            event_camera_pivot)
                        st.markdown(event_heatmap_html, unsafe_allow_html=True)
                    else:
                        st.info(
                            "No event data available for the selected period.")
                else:
                    st.info("No matching data found for top events.")
            else:
                st.info("Events column not found in the data.")
        else:
            st.info("No camera/area data available for the selected period.")

        # Area wise Report For Top 3 Areas Section
        st.markdown("<br><br>", unsafe_allow_html=True)

        # Determine which column to use for areas
        area_column = None
        if 'Area' in df_selected.columns:
            area_column = 'Area'
        elif 'Camera' in df_selected.columns:
            area_column = 'Camera'

        # Area Wide Summary Slide (before the main heading)
        if area_column and len(df_selected) > 0:
            # Create two columns for the summary layout
            col_left, col_right = st.columns(2)

            with col_left:
                st.markdown("### 🏢 All Areas Summary")

                # Generate pie chart for all areas contribution
                area_counts = df_selected[area_column].value_counts()
                total_events = len(df_selected)

                # Create pie chart
                area_pie_img = plot_events_per_area_pie(df_selected)
                st.image(
                    f"data:image/png;base64,{area_pie_img}",
                    use_container_width=True)
            with col_right:
                st.markdown("#### 🏆 Top 3 Areas")

                # Get top 3 areas by event count
                top_areas = df_selected[area_column].value_counts().head(3)

                # Display top 3 areas with enhanced styling
                for rank, (area_name, area_count) in enumerate(top_areas.items(), 1):
                    area_percentage = (area_count / total_events * 100) if total_events > 0 else 0

                    # Calculate area statistics
                    area_data = df_selected[df_selected[area_column] == area_name]
                    if "Events" in area_data.columns:
                        near_miss_pattern = r"near[-s]*miss"
                        area_near_miss = area_data[area_data["Events"].str.contains(
                            near_miss_pattern, case=False, na=False, regex=True)]
                        area_near_miss_count = len(area_near_miss)
                    else:
                        area_near_miss_count = 0

                    # Rank styling
                    rank_colors = {1: "#ffd700", 2: "#c0c0c0", 3: "#cd7f32"}  # Gold, Silver, Bronze
                    rank_color = rank_colors.get(rank, "#gray")

                    st.markdown(f"""
                    <div style="background-color: white; border: 2px solid {rank_color}; border-radius: 10px; padding: 15px; margin: 10px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                        <div style="display: flex; align-items: center; margin-bottom: 10px;">
                            <div style="background-color: {rank_color}; color: white; border-radius: 50%; width: 30px; height: 30px; display: flex; align-items: center; justify-content: center; font-weight: bold; margin-right: 10px;">
                                #{rank}
                            </div>
                            <h4 style="margin: 0; color: #374151;">{area_name}</h4>
                        </div>
                        <div style="color: #6b7280; font-size: 14px;">
                            <strong>📊 Total Events:</strong> {area_count}<br>
                            <strong>📈 Percentage:</strong> {area_percentage:.1f}%<br>
                            <strong>⚠️ Near Miss:</strong> {area_near_miss_count}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---", unsafe_allow_html=True)  # Add a separator
            st.markdown("<br>", unsafe_allow_html=True)

            # Detailed Area-wise Report Section
            st.markdown("### 📊 Detailed Area Analysis")
            st.markdown("*In-depth analysis of top 3 performing areas*")
            st.markdown("<br>", unsafe_allow_html=True)

            # Get top 3 areas by event count (reuse the previous calculation)
            for area_name, area_count in top_areas.items():
                # Filter data for this specific area
                area_data = df_selected[df_selected[area_column] == area_name]

                # Calculate area statistics
                area_total_events = len(area_data)
                area_percentage = (area_total_events / len(df_selected) * 100) if len(df_selected) > 0 else 0

                # Calculate near miss events for this area
                if 'Events' in area_data.columns:
                    near_miss_pattern = r'near[-\s]*miss'
                    area_near_miss = area_data[area_data['Events'].str.contains(
                        near_miss_pattern, case=False, na=False, regex=True)]
                    area_near_miss_count = len(area_near_miss)
                else:
                    area_near_miss_count = 0

                # Create area heading
                st.markdown(f"### Area: {area_name}", unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

                # Create two columns for the metrics
                metric_col1, metric_col2 = st.columns(2)

                with metric_col1:
                    # Critical Near Miss Events section
                    st.markdown(f"""
                    <div style="background-color: #dcfce7; border: 1px solid #bbf7d0; border-radius: 8px; padding: 15px; margin: 10px 0;">
                        <h4 style="margin: 0 0 5px 0; color: #166534;">Critical Near Miss Events</h4>
                        <p style="margin: 0; color: #166534;">Near Miss Events: {area_near_miss_count}</p>
                    </div>
                    """, unsafe_allow_html=True)

                with metric_col2:
                    # High Impact Analysis section
                    st.markdown(f"""
                    <div style="background-color: #dbeafe; border: 1px solid #bfdbfe; border-radius: 8px; padding: 15px; margin: 10px 0;">
                        <h4 style="margin: 0 0 5px 0; color: #1e40af;">High Impact Analysis</h4>
                        <p style="margin: 0; color: #1e40af;">Total Events: {area_total_events} ({area_percentage:.1f}% of total)</p>
                    </div>
                    """, unsafe_allow_html=True)

                # Distribution Charts section
                st.markdown("#### 📊 Distribution Analysis")

                # Generate pie charts
                camera_events_pie_img = plot_camera_events_pie(area_data, area_name, area_column)
                camera_severity_pie_img = plot_camera_severity_pie(area_data, area_name, area_column)

                # Display pie charts side by side
                if camera_events_pie_img or camera_severity_pie_img:
                    pie_col1, pie_col2 = st.columns(2)

                    with pie_col1:
                        if camera_events_pie_img:
                            st.markdown("**📈 Camera vs Events Distribution**")
                            st.image(f"data:image/png;base64,{camera_events_pie_img}", use_container_width=True)
                        else:
                            st.info("No camera vs events data available for pie chart.")

                    with pie_col2:
                        if camera_severity_pie_img:
                            st.markdown("**⚠️ Camera vs Severity Distribution**")
                            st.image(f"data:image/png;base64,{camera_severity_pie_img}", use_container_width=True)
                        else:
                            st.info("No camera vs severity data available for pie chart.")

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown("---", unsafe_allow_html=True)  # Add separator between areas
        else:
            st.info("No area data available for the selected period.")

        # Top Videos by Priority Section (inserted between Top 3 Areas and Analysis sections)
        st.markdown("<br><br>", unsafe_allow_html=True)
        if has_video_paths:
            display_latest_videos_section(df_selected, uploaded_file)
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---", unsafe_allow_html=True)  # Add separator
            st.markdown("<br>", unsafe_allow_html=True)



        # Show the 4 main charts as per user requirement
        st.header(f"Analysis for {date_range_text}")

        # Generate all chart images
        events_by_type_bar_img = plot_events_by_type_bar(df_selected)
        events_by_area_bar_img = plot_events_by_area_bar(df_selected)
        events_by_severity_pie_img = plot_events_by_severity_pie(df_selected)
        events_by_type_pie_img = plot_events_by_type_pie(df_selected)

        # Display charts in 2x2 grid layout
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📊 Events by Type")
            st.image(
                f"data:image/png;base64,{events_by_type_bar_img}",
                use_container_width=True)

        with col2:
            st.subheader("🏢 Events by Area")
            st.image(
                f"data:image/png;base64,{events_by_area_bar_img}",
                use_container_width=True)

        col3, col4 = st.columns(2)

        with col3:
            st.subheader("⚠️ Events by Severity")
            st.image(
                f"data:image/png;base64,{events_by_severity_pie_img}",
                use_container_width=True)

        with col4:
            st.subheader("📊 Events by Type (Pie Chart)")
            st.image(
                f"data:image/png;base64,{events_by_type_pie_img}",
                use_container_width=True)

        # Generate and show trend graphs
        # Create appropriate variables for plotting functions based on date
        # range
        start_month = start_date.month
        start_year = start_date.year
        end_month = end_date.month
        end_year = end_date.year

        # For title purposes, use start date info or create a range description
        if start_month == end_month and start_year == end_year:
            month_names = ['', 'January', 'February', 'March', 'April', 'May', 'June',
                           'July', 'August', 'September', 'October', 'November', 'December']
            month_name = month_names[start_month]
            year = start_year
        else:
            month_name = f"{start_date.strftime('%b')} - {end_date.strftime('%b')}"
            year = f"{start_year}" if start_year == end_year else f"{start_year}-{end_year}"

        events_img = plot_events_per_day_range(
            df_selected, start_date, end_date, date_range_text)
        interval_img = plot_avg_interval_range(
            df_selected, start_date, end_date, date_range_text)

        st.header("Trend Graphs")
        st.image(
            f"data:image/png;base64,{events_img}",
            use_container_width=True)
        st.image(
            f"data:image/png;base64,{interval_img}",
            use_container_width=True)

        # Generate and show weekly/monthly trends based on selected date range
        # (filtered data)
        st.header(f"Weekly and Monthly Trends for {date_range_text}")

        # Generate weekly trend
        weekly_trend_img = plot_weekly_trend(df_selected)
        if weekly_trend_img:
            st.subheader("📊 Weekly Trend Analysis")
            st.image(
                f"data:image/png;base64,{weekly_trend_img}",
                use_container_width=True)

        # Generate monthly trend
        monthly_trend_img = plot_monthly_trend(df_selected)
        if monthly_trend_img:
            st.subheader("📈 Monthly Trend Analysis")
            st.image(
                f"data:image/png;base64,{monthly_trend_img}",
                use_container_width=True)

        # Show info if no trends generated
        if not weekly_trend_img and not monthly_trend_img:
            st.info("📌 Weekly and monthly trends require data spanning multiple weeks/months. Current dataset appears to cover a shorter time period.")
        elif not weekly_trend_img:
            st.info("📌 Weekly trend not shown - dataset covers less than one week.")
        elif not monthly_trend_img:
            st.info("📌 Monthly trend not shown - dataset covers only one month.")

        
        # PowerPoint Report Generation Section
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.header("📊 Professional PowerPoint Report Generation")
        st.markdown(
            "*Generate industry-standard, executive-ready PowerPoint presentations with advanced analytics*")

        # PowerPoint generation options
        col1, col2 = st.columns([3, 1])

        with col1:
            # Report title input
            report_title = st.text_input(
                "Report Title",
                value="Event Analytics Report",
                help="Enter the title that will appear on the first slide of your report"
            )

            # PowerPoint file name input
            file_name = st.text_input(
                "Report File Name",
                value="Event_Analytics_Report",
                help="Specify the name for your PowerPoint file (without .pptx extension). The current date and time will be added automatically."
            )

            # Custom logo upload
            uploaded_logo = st.file_uploader(
                "Upload Custom Logo (Optional)",
                type=['png', 'jpg', 'jpeg'],
                help="Upload your company logo to replace VISIONIFY branding. Logo will be positioned in the top right corner of each slide."
            )

        with col2:
            st.markdown("<br>", unsafe_allow_html=True)

            # Show comparison data status
            if hasattr(st.session_state, 'comparison_data') and st.session_state.comparison_data:
                st.success("✅ Excel comparison data available - will be included in PowerPoint")
            else:
                st.info("ℹ️ Upload Excel files above to include comparison analysis in PowerPoint")

            # Generate PowerPoint button
            if st.button(
                "📊 Generate Professional Event Analytics (Up to 10 Slides)",
                type="primary",
                use_container_width=True
            ):
                with st.spinner("🔄 Creating comprehensive analytics presentation with 10 slides..."):
                    try:
                        # Handle custom logo if uploaded
                        custom_logo_path = None
                        if uploaded_logo is not None:
                            # Save uploaded logo temporarily (AWS compatible)
                            import uuid
                            file_extension = uploaded_logo.name.split('.')[-1]
                            temp_logo_name = f"temp_logo_{uuid.uuid4().hex[:8]}.{file_extension}"
                            custom_logo_path = os.path.join(os.getcwd(), temp_logo_name)
                            with open(custom_logo_path, 'wb') as tmp_file:
                                tmp_file.write(uploaded_logo.read())

                        # Generate the PowerPoint report using accurate
                        # centralized functions
                        kpis = get_accurate_kpis(df_selected)

                        # Get comparison data from session state if available
                        comparison_data = None
                        if hasattr(st.session_state, 'comparison_data') and st.session_state.comparison_data:
                            comp_data = st.session_state.comparison_data
                            comparison_data = [
                                comp_data['df1'],
                                comp_data['df2'],
                                comp_data['label1'],
                                comp_data['label2'],
                                comp_data['metrics']
                            ]

                        ppt_buffer = create_powerpoint_report(
                            df_selected, report_title, custom_logo_path, comparison_data)

                        # Create download button with custom file name
                        current_date = datetime.now().strftime("%Y%m%d_%H%M%S")
                        # Clean file name to remove invalid characters
                        clean_file_name = "".join(
                            c for c in file_name if c.isalnum() or c in (
                                ' ', '-', '_')).rstrip()
                        if not clean_file_name:  # Fallback if name is empty
                            clean_file_name = "Event_Analytics_Report"
                        filename = f"{clean_file_name}_{current_date}.pptx"
                        st.success("✅ Professional Event Analytics with 10 slides generated successfully!")
                        st.download_button(
                            label="📥 Download Professional Event Analytics",
                            data=ppt_buffer.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                        # Clean up temporary logo file
                        if custom_logo_path and os.path.exists(
                                custom_logo_path):
                            try:
                                os.unlink(custom_logo_path)
                            except BaseException:
                                pass  # Ignore cleanup errors

                        # Show report summary
                        st.markdown("### 📋 Report Summary")
                        col_a, col_b, col_c = st.columns(3)

                        with col_a:
                            st.metric(
                                "Total Events", f"{kpis['total_events']:,}")
                        with col_b:
                            st.metric(
                                "Event Types", kpis.get(
                                    'total_event_types', 'N/A'))
                        with col_c:
                            st.metric(
                                "Areas Monitored", kpis.get(
                                    'total_areas', 'N/A'))

                        st.info(f"🎯 **Report covers:** {kpis['date_range']}")

                        # Usage instructions
                        with st.expander("Report Analytics Guide", expanded=False):
                            st.markdown("""
                            <h3>Report Content (10 Slides Total):</h3>

                            <h4>Slide 1: Event Trends Dashboard</h4>
                            <ul>
                                <li><strong>Event Trends Title</strong> - Clean header with VISIONIFY branding</li>
                                <li><strong>4 KPI Cards</strong> - Total Events, Event Types, Areas Monitored, Review Percent</li>
                                <li><strong>Events by Type Chart</strong> - Clean bar chart with event counts</li>
                            </ul>

                            <h4>Slide 2: High Impact Analysis</h4>
                            <ul>
                                <li><strong>High Impact Analysis Title</strong> - Professional header with VISIONIFY branding</li>
                                <li><strong>3 Impact Cards</strong> - Near-Miss (green), Emergency (red), High Risk (red) events</li>
                                <li><strong>Top 3 Contributing Cameras</strong> - Detailed camera cards with rankings and severity counts</li>
                            </ul>

                            <h4>Slide 3: Top 10 Cameras By Event Count</h4>
                            <ul>
                                <li><strong>Professional Heatmap Table</strong> - Color-coded event counts by camera and event type</li>
                                <li><strong>Top 10 Cameras</strong> - Ranked by total event count with accurate data</li>
                                <li><strong>Top 5 Event Types</strong> - Most frequent events as columns</li>
                                <li><strong>Color-coded Cells</strong> - Red intensity based on event frequency</li>
                            </ul>

                            <h4>Slide 4: Top 5 Event Types</h4>
                            <ul>
                                <li><strong>Professional Heatmap Table</strong> - Color-coded event counts by event type and camera</li>
                                <li><strong>Top 5 Event Types</strong> - Most frequent events as rows</li>
                                <li><strong>Top 5 Cameras</strong> - Highest event cameras as columns</li>
                                <li><strong>Color-coded Cells</strong> - Red intensity showing event distribution patterns</li>
                            </ul>

                            <h4>Slide 5: Events Analysis</h4>
                            <ul>
                                <li><strong>Side-by-Side Bar Charts</strong> - Professional dual-chart layout</li>
                                <li><strong>Events by Type</strong> - Blue bars showing event type distribution with percentages</li>
                                <li><strong>Events by Area</strong> - Green bars showing area-wise event distribution with percentages</li>
                                <li><strong>Professional Styling</strong> - Clean charts with value labels and grid lines</li>
                            </ul>

                            <h4>Slide 6: Recommendations</h4>
                            <ul>
                                <li><strong>Interactive Template</strong> - Structured layout for manual recommendations entry</li>
                                <li><strong>Priority-Based Sections</strong> - High (red), Medium (yellow), General (green) recommendations</li>
                                <li><strong>Professional Styling</strong> - Color-coded boxes with clear formatting guidelines</li>
                                <li><strong>User-Friendly</strong> - Click-to-edit placeholders and helpful instructions</li>
                            </ul>

                            <h4>Default Company Slides (Automatically Added):</h4>
                            <ul>
                                <li><strong>Slide 7:</strong> Floor Dashboard Recommendation</li>
                                <li><strong>Slide 8:</strong> Additional Content Slide</li>
                                <li><strong>Slide 9:</strong> Additional Content Slide</li>
                                <li><strong>Slide 10:</strong> Supported AI Scenarios Overview</li>
                                <li><strong>Slide 11:</strong> Contact Information</li>
                            </ul>

                            <h3>Report Features:</h3>
                            <ul>
                                <li><strong>Real-time Data</strong> - Based on your current filtered dataset</li>
                                <li><strong>Professional Design</strong> - Clean layouts with VISIONIFY branding</li>
                                <li><strong>High Quality</strong> - 300 DPI charts for crisp presentation</li>
                                <li><strong>Comprehensive Analytics</strong> - Impact analysis, heatmaps, and area breakdowns</li>
                                <li><strong>Interactive Recommendations</strong> - Structured template for insights</li>
                                <li><strong>Company Slides</strong> - Floor dashboard and AI scenarios included</li>
                            </ul>

                            <h3>Next Steps:</h3>
                            <ol>
                                <li><strong>Download</strong> your comprehensive Event Analytics Report (10 slides)</li>
                                <li><strong>Edit</strong> the Recommendations slide with your specific insights</li>
                                <li><strong>Review</strong> all analytics and company information slides</li>
                                <li><strong>Present</strong> your professional report with executive-ready insights</li>
                            </ol>
                            """, unsafe_allow_html=True)

                        st.markdown("<h3>Quick Actions</h3>", unsafe_allow_html=True)
                        st.markdown("""
                        <ul>
                            <li><strong>Download</strong> your comprehensive Event Analytics Report (10 slides)</li>
                            <li><strong>Review</strong> analytics slides: Event Trends + High Impact Analysis + Top 10 Cameras + Top 5 Event Types + Events Analysis + Recommendations</li>
                            <li><strong>Customize</strong> the Recommendations slide with your specific insights and action items</li>
                            <li><strong>Present</strong> with included company slides: Floor Dashboard + AI Scenarios + Contact Info</li>
                        </ul>
                        """, unsafe_allow_html=True)

                        st.markdown("- Ensure your data has the required columns (Date, Events)")
                        st.markdown("- Check that the selected date range contains valid data")
                        st.markdown("- Try refreshing the page and uploading the file again")
                    except Exception as e:
                        st.error(f"Error generating PowerPoint report: {str(e)}")
                        st.markdown("Please try the following:")
                        st.markdown("- Ensure your data has the required columns (Date, Events)")
                        st.markdown("- Check that the selected date range contains valid data")
                        st.markdown("- Try refreshing the page and uploading the file again")


def extract_video_links(df, top_n=5, excel_file=None):
    """
    Extract the latest video links for each event type, prioritized by severity and recency.
    Now supports extracting actual hyperlinks from Excel files.
    """
    try:
        # Check if we have the required columns
        if 'Events' not in df.columns or 'Date' not in df.columns:
            return {}
        
        # Check if we have video path column
        video_columns = ['Event Video Path', 'Video Path', 'Video', 'Path']
        video_column = None
        for col in video_columns:
            if col in df.columns:
                video_column = col
                break
        
        if not video_column:
            return {}
        
        # Create a copy of the dataframe with only the columns we need
        video_df = df[['Events', 'Date', video_column]].copy()
        
        # Add Camera and Area columns if available
        if 'Camera' in df.columns:
            video_df['Camera'] = df['Camera']
        if 'Area' in df.columns:
            video_df['Area'] = df['Area']
        if 'Severity' in df.columns:
            video_df['Severity'] = df['Severity']
        
        # Remove rows with missing video paths
        video_df = video_df.dropna(subset=[video_column])
        video_df = video_df[video_df[video_column] != '']
        
        if video_df.empty:
            return {}
        
        # Convert Date column to datetime if it's not already
            video_df['Date'] = pd.to_datetime(video_df['Date'], dayfirst=True, errors='coerce')
        video_df = video_df.dropna(subset=['Date'])
        
        if video_df.empty:
            return {}
        
        # Extract hyperlinks from Excel file if provided
        hyperlinks = {}
        if excel_file is not None:
            try:
                hyperlinks = extract_hyperlinks_from_excel(excel_file, video_column)
            except Exception as e:
                hyperlinks = {}
        
        # Create severity ranking (higher number = higher priority)
        severity_rank = {
            'Critical': 4,
            'High': 3,
            'Moderate': 2,
            'Medium': 2,  # Alternative spelling
            'Low': 1,
            'Unknown': 0
        }
        
        # Add severity rank and normalize severity values
        video_df['severity_rank'] = video_df['Severity'].fillna('Unknown').astype(str).map(
            lambda x: severity_rank.get(x, 0)
        )
        
        # Get latest video for each event type first
        event_videos = {}
        for event_type in video_df['Events'].unique():
            try:
                event_videos_subset = video_df[video_df['Events'] == event_type]
                
                # Check if we have any videos for this event type after filtering
                if event_videos_subset.empty:
                    continue
                    
                # Check if we have any valid dates for this event type
                valid_dates = event_videos_subset['Date'].dropna()
                if valid_dates.empty:
                    continue
                    
                # Get the most recent video for this event type
                latest_video = event_videos_subset.loc[event_videos_subset['Date'].idxmax()]
                
                # Ensure we have the required video path
                if pd.isna(latest_video[video_column]) or latest_video[video_column] == '':
                    continue
                
                # Get the row index in the original dataframe to match with hyperlinks
                original_row_index = latest_video.name
                
                # Use hyperlink if available, otherwise fall back to display text
                video_path = latest_video[video_column]
                if original_row_index in hyperlinks:
                    actual_hyperlink = hyperlinks[original_row_index]
                    # Use the actual hyperlink URL instead of display text
                    video_path = actual_hyperlink
                
                event_videos[event_type] = {
                    'video_path': video_path,
                    'display_text': latest_video[video_column],  # Keep original display text
                    'date': latest_video['Date'],
                    'camera': latest_video.get('Camera', 'Unknown'),
                    'area': latest_video.get('Area', 'Unknown'),
                    'severity': latest_video.get('Severity', 'Unknown'),
                    'severity_rank': latest_video.get('severity_rank', 0),
                    'event_type': event_type,
                    'has_hyperlink': original_row_index in hyperlinks
                }
            except Exception as e:
                # Skip this event type if there's any error processing it
                print(f"Warning: Error processing event type '{event_type}': {e}")
                continue
        
        if not event_videos:
            return {}
        
        # Sort videos by severity rank (descending) and then by date (descending)
        sorted_videos = sorted(
            event_videos.items(),
            key=lambda x: (x[1]['severity_rank'], x[1]['date']),
            reverse=True
        )
        
        # Return only top N videos
        top_videos = {}
        for i, (event_type, video_info) in enumerate(sorted_videos[:top_n]):
            top_videos[event_type] = video_info
        
        return top_videos
    
    except Exception as e:
        # Return empty dict if there's any global error
        print(f"Warning: Error extracting video links: {e}")
        return {}


def create_chart_for_ppt(df, chart_type, title, save_path):
    """
    Create high-quality charts optimized for PowerPoint presentations
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('white')

    if chart_type == 'events_by_type':
        # Bar chart for events by type
        if 'Events' in df.columns:
            event_counts = df['Events'].value_counts().head(10)
            bars = ax.bar(event_counts.index, event_counts.values, color=PROFESSIONAL_BLUE, alpha=0.8)

            # Add value labels on bars
            for bar, count in zip(bars, event_counts.values):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                        str(count), ha='center', va='bottom', fontweight='bold', fontsize=10)

            ax.set_title(title, fontsize=16, fontweight='bold', pad=20)
            ax.set_xlabel('Event Type', fontsize=12, fontweight='bold')
            ax.set_ylabel('Number of Events', fontsize=12, fontweight='bold')

            # Wrap long labels
            wrapped_labels = ['\n'.join(textwrap.wrap(label, width=15)) for label in event_counts.index]
            ax.set_xticks(range(len(event_counts)))
            ax.set_xticklabels(wrapped_labels, rotation=45, ha='right', fontsize=9)

    elif chart_type == 'events_by_area':
        # Bar chart for events by area
        area_col = 'Area' if 'Area' in df.columns else 'Camera'
        if area_col in df.columns:
            area_counts = df[area_col].value_counts().head(10)
            bars = ax.bar(area_counts.index, area_counts.values, color=PROFESSIONAL_BLUE, alpha=0.8)

            # Add value labels on bars
            for bar, count in zip(bars, area_counts.values):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                        str(count), ha='center', va='bottom', fontweight='bold', fontsize=10)

            ax.set_title(title, fontsize=16, fontweight='bold', pad=20)
            ax.set_xlabel('Area/Location', fontsize=12, fontweight='bold')
            ax.set_ylabel('Number of Events', fontsize=12, fontweight='bold')

            # Wrap long labels
            wrapped_labels = ['\n'.join(textwrap.wrap(str(label), width=12)) for label in area_counts.index]
            ax.set_xticks(range(len(area_counts)))
            ax.set_xticklabels(wrapped_labels, rotation=45, ha='right', fontsize=9)

    elif chart_type == 'events_by_severity':
        # Pie chart for events by severity
        if 'Severity' in df.columns:
            severity_counts = df['Severity'].value_counts()
            colors = PIE_COLORS[:len(severity_counts)]
            if len(severity_counts) > len(PIE_COLORS):
                colors.extend(['#94A3B8'] * (len(severity_counts) - len(PIE_COLORS)))
            wedges, texts, autotexts = ax.pie(severity_counts.values, labels=severity_counts.index,
                                              autopct='%1.1f%%', startangle=90, colors=colors)

            # Enhance text formatting
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(11)

            for text in texts:
                text.set_fontsize(12)
                text.set_fontweight('bold')

            ax.set_title(title, fontsize=16, fontweight='bold', pad=20)
        else:
            ax.text(0.5, 0.5, 'No Severity data available', ha='center', va='center',
                    transform=ax.transAxes, fontsize=14)
            ax.set_title(title, fontsize=16, fontweight='bold', pad=20)

    elif chart_type == 'monthly_trend':
        # Line chart for monthly trends
        if 'Date' in df.columns:
            df_copy = df.copy()
            df_copy['Month'] = pd.to_datetime(df_copy['Date']).dt.to_period('M')
            monthly_counts = df_copy['Month'].value_counts().sort_index()

            ax.plot(range(len(monthly_counts)), monthly_counts.values,
                    marker='o', linewidth=3, markersize=8, color='#2E86AB')
            ax.fill_between(range(len(monthly_counts)), monthly_counts.values, alpha=0.3, color='#2E86AB')

            # Add value labels
            for i, count in enumerate(monthly_counts.values):
                ax.text(i, count + max(monthly_counts.values) * 0.02, str(count),
                        ha='center', va='bottom', fontweight='bold', fontsize=10)

            ax.set_title(title, fontsize=16, fontweight='bold', pad=20)
            ax.set_xlabel('Month', fontsize=12, fontweight='bold')
            ax.set_ylabel('Number of Events', fontsize=12, fontweight='bold')
            ax.set_xticks(range(len(monthly_counts)))
            ax.set_xticklabels([str(month) for month in monthly_counts.index], rotation=45)

    # Common formatting
    ax.grid(True, alpha=0.3)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()

    # Save with high DPI for PowerPoint
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def calculate_kpis(df):
    """
    Calculate key performance indicators from the data
    """
    kpis = {}

    # Basic metrics
    kpis['total_events'] = len(df)
    kpis['date_range'] = f"{df['Date'].min().strftime('%B %Y')} - {df['Date'].max().strftime('%B %Y')}"

    # Monthly breakdown
    if 'Date' in df.columns:
        df_copy = df.copy()
        df_copy['Month'] = pd.to_datetime(df_copy['Date']).dt.to_period('M')
        monthly_counts = df_copy['Month'].value_counts().sort_index()
        kpis['monthly_breakdown'] = {str(month): count for month, count in monthly_counts.items()}

    # Event types
    if 'Events' in df.columns:
        event_counts = df['Events'].value_counts()
        kpis['top_event_types'] = dict(event_counts.head(5))
        kpis['total_event_types'] = len(event_counts)

    # Severity analysis
    if 'Severity' in df.columns:
        severity_counts = df['Severity'].value_counts()
        kpis['severity_breakdown'] = dict(severity_counts)

        # Calculate high-risk events
        high_risk_patterns = ['critical', 'high', 'severe', 'major']
        high_risk_count = 0
        for pattern in high_risk_patterns:
            high_risk_count += df['Severity'].str.contains(pattern, case=False, na=False).sum()
        kpis['high_risk_events'] = high_risk_count
        kpis['high_risk_percentage'] = (high_risk_count / len(df) * 100) if len(df) > 0 else 0

    # Area analysis
    area_col = 'Area' if 'Area' in df.columns else 'Camera'
    if area_col in df.columns:
        area_counts = df[area_col].value_counts()
        kpis['top_areas'] = dict(area_counts.head(5))
        kpis['total_areas'] = len(area_counts)

    # Near miss analysis
    if 'Events' in df.columns:
        near_miss_pattern = r'near[-\s]*miss'
        near_miss_count = df['Events'].str.contains(near_miss_pattern, case=False, na=False, regex=True).sum()
        kpis['near_miss_events'] = near_miss_count
        kpis['near_miss_percentage'] = (near_miss_count / len(df) * 100) if len(df) > 0 else 0

    return kpis

    try:
        fig, ax = plt.subplots(figsize=(2, 1))
        ax.text(0.5, 0.5, 'VISIONIFY', ha='center', va='center',
                fontsize=14, fontweight='bold', color='#1f2937')
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis('off')
        fig.patch.set_facecolor('white')
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close()
        return save_path
    except Exception as e:
        print(f"Error creating logo: {e}")
        return None


def add_logo_to_slide(slide, custom_logo_path=None, x_position=9.0, y_position=0.3, width=1.5, height=0.75):
    """
    Add a logo to the slide with improved dimensions and positioning at top right
    Args:
        slide: PowerPoint slide object
        custom_logo_path: Path to custom logo (optional)
        x_position: Horizontal position in inches from left (9.0 for top right)
        y_position: Vertical position in inches from top
        width: Width of the logo in inches
        height: Height of the logo in inches
    """
    try:
        if custom_logo_path and os.path.exists(custom_logo_path):
            # Use custom logo if provided
            logo_path = custom_logo_path
        else:
            # Create and use default Visionify logo
            logo_path = "visionify.jpeg"
            if not os.path.exists(logo_path):
                create_visionify_logo(logo_path)

        # Calculate the slide width to properly position the logo
        slide_width = 13.33  # Standard PowerPoint slide width in inches

        # Adjust x_position to be relative to right edge
        actual_x = slide_width - width - 0.5  # 0.5 inch margin from right

        # Add picture while maintaining aspect ratio
        pic = slide.shapes.add_picture(
            logo_path,
            Inches(actual_x),
            Inches(y_position),
            width=Inches(width),
            height=Inches(height)
        )

        return True
    except Exception as e:
        print(f"Error adding logo to slide: {str(e)}")
        return False


def add_custom_logo_to_slide(slide, logo_path, x_position=10.5, y_position=0.2, width=1.5, height=1.0):
    """
    Legacy function - now calls the new add_logo_to_slide function
    """
    add_logo_to_slide(slide, logo_path, x_position, y_position, width, height)


def add_default_slides(prs, blank_slide_layout):
    """
    Add default slides from image files in sequence without empty slides:
    slide1.png  -> first slide
    slide8.png  -> after the analytics slides
    slide9.png  -> after slide8
    slide10.png -> after slide9
    """
    # Define the default slides in sequence
    default_slides = [
        'slide1.png',    # Visionify AI Monitoring Solution (always first) 
        'slide8.png',   # Floor Dashboard recommendation (after analytics)
        'slide9.png',   # Supported AI Scenarios
        'slide10.png'    # Contact information
    ]

    slides_added = 0

    # Handle slide1.png separately as it needs to be the first slide
    if os.path.exists('slide1.png'):
        try:
            # If we already have slides, insert at position 0
            if len(prs.slides) > 0:
                slide = prs.slides.add_slide(blank_slide_layout)
                xml_slides = prs.slides._sldIdLst
                slides = list(xml_slides)
                slide_element = slides[-1]
                xml_slides.remove(slide_element)
                xml_slides.insert(0, slide_element)
            else:
                # If no slides exist, just add it
                slide = prs.slides.add_slide(blank_slide_layout)

            # Set slide background to white
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 255, 255)

            # Add the image to fill the entire slide (16:9 ratio)
            slide.shapes.add_picture('slide1.png', Inches(0), Inches(0),
                                     width=Inches(13.33), height=Inches(7.5))
            slides_added += 1
            print("Added slide1.png as first slide")
        except Exception as e:
            print(f"Error adding slide1.png: {e}")

    # Add the remaining slides in sequence (8, 9, 12, 13, 14)
    for filename in default_slides[1:]:  # Skip slide1.png as it's already handled
        if os.path.exists(filename):
            try:
                # Add the slide at the end
                slide = prs.slides.add_slide(blank_slide_layout)

                # Set slide background to white
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(255, 255, 255)

                # Add the image to fill the entire slide (16:9 ratio)
                slide.shapes.add_picture(filename, Inches(0), Inches(0),
                                         width=Inches(13.33), height=Inches(7.5))

                slides_added += 1
                print(f"Added {filename} to presentation")

            except Exception as e:
                print(f"Error adding {filename}: {e}")
        else:
            print(f"Default slide not found: {filename}")

    if slides_added > 0:
        print(f"Successfully added {slides_added} default slide(s)")
    else:
        print("No default slides were added (files not found)")

    return slides_added


def create_powerpoint_report(
    df_selected,
    report_title="Visionify Weekly Report",
    custom_logo_path=None,
     comparison_data=None):
    """
    Generate a basic PowerPoint presentation structure
    Ready for slide-by-slide implementation based on screenshots
    """
    # Create presentation
    prs = Presentation()

    # Set slide size to widescreen (16:9)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # Set up slide layouts
    title_slide_layout = prs.slide_layouts[0]  # Title slide
    content_slide_layout = prs.slide_layouts[1]  # Title and content
    blank_slide_layout = prs.slide_layouts[6]  # Blank slide

    # Calculate KPIs and data analysis using filtered data (same as dashboard)
    kpis = get_accurate_kpis(df_selected)
    event_data = get_accurate_event_counts(df_selected)
    camera_data = get_accurate_camera_data(df_selected)
    severity_data = get_accurate_severity_data(df_selected)

    # Calculate camera_contributions exactly like the dashboard does using df_selected
    camera_contributions = []

    # Determine which column to use for cameras (same logic as dashboard)
    camera_column = None
    if 'Camera' in df_selected.columns:
        camera_column = 'Camera'
    elif 'Area' in df_selected.columns:
        camera_column = 'Area'

    if camera_column:
        camera_counts = df_selected[camera_column].value_counts().head(3)

        for camera, count in camera_counts.items():
            # Calculate high severity events for this camera with flexible matching
            if 'Severity' in df_selected.columns:
                camera_df = df_selected[df_selected[camera_column] == camera]
                severity_str = camera_df['Severity'].astype(str).str.lower()
                high_severity_pattern = r'(critical|high|severe|major)'
                camera_high_severity = severity_str.str.contains(
                    high_severity_pattern, case=False, na=False, regex=True).sum()
            else:
                camera_high_severity = 0

            # Extract location/plant info from camera name if available
            location = "Unknown Location"
            if isinstance(camera, str):
                if 'plant' in camera.lower():
                    # Extract plant info (e.g., "Plant-2 Bagging Station" -> "Plant-2 Bagging")
                    import re
                    plant_match = re.search(
                        r'plant[-\s]*\d+[-\s]*\w*', camera, re.IGNORECASE)
                    if plant_match:
                        location = plant_match.group(0)
                    else:
                        location = "Plant Location"
                elif any(keyword in camera.lower() for keyword in ['area', 'zone', 'section', 'dept']):
                    location = "Operational Area"
                else:
                    location = "Monitoring Point"

            camera_contributions.append({
                'camera': camera,
                'total': count,
                'high_severity': camera_high_severity,
                'location': location
            })

    # Use camera_contributions for PowerPoint (same as dashboard)
    camera_severity_data = camera_contributions

    # Get date range for report using filtered data
    date_range = f"{df_selected['Date'].min().strftime('%d-%b-%Y')} to {df_selected['Date'].max().strftime('%d-%b-%Y')}"

    # Create temporary directory for charts (AWS compatible)
    temp_dir = create_safe_temp_dir()

    try:
        # SLIDE 1: Event Trends Dashboard (matching the screenshot exactly)
        slide1 = prs.slides.add_slide(blank_slide_layout)

        # Set slide background to white
        background = slide1.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)

        # Add main title "Event Trends" in top left
        main_title = slide1.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(6), Inches(0.8))
        main_title_frame = main_title.text_frame
        main_title_frame.text = "Event Trends"
        main_title_frame.paragraphs[0].font.size = Pt(36)
        main_title_frame.paragraphs[0].font.bold = True
        main_title_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark gray
        main_title_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        main_title_frame.paragraphs[0].font.name = "Arial"

        # Add logo (VISIONIFY by default or custom if provided)
        add_logo_to_slide(slide1, custom_logo_path)

        # Create KPI metric cards (4 cards in a row) with rounded corners - using accurate data
        kpi_metrics = [
            ("📈", "Total Events", str(kpis['total_events']), 1.0),
            ("⚠️", "Event Types", str(kpis['total_event_types']), 4.0),
            ("⭕", "Areas Monitored", str(kpis['total_areas']), 7.0),
            ("✅", "Review Percent", str(kpis['review_percent']) + "%", 10.0)
        ]

        for icon, label, value, x_pos in kpi_metrics:
            # Create KPI card with light gray background and rounded corners
            kpi_card = slide1.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x_pos),
                Inches(1.3),
                Inches(2.7),
                Inches(1.2))
            kpi_card.fill.solid()
            kpi_card.fill.fore_color.rgb = RGBColor(248, 250, 252)  # Very light gray
            kpi_card.line.color.rgb = RGBColor(229, 231, 235)  # Light border
            kpi_card.line.width = Pt(1)

            # Add icon and label text (smaller, gray)
            icon_label_box = slide1.shapes.add_textbox(Inches(x_pos + 0.2), Inches(1.5), Inches(2.3), Inches(0.4))
            icon_label_frame = icon_label_box.text_frame
            icon_label_frame.text = f"{icon} {label}"
            icon_label_frame.paragraphs[0].font.size = Pt(11)
            icon_label_frame.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)  # Gray text
            icon_label_frame.paragraphs[0].font.name = "Arial"
            icon_label_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add value text (large, bold, black)
            value_box = slide1.shapes.add_textbox(Inches(x_pos + 0.2), Inches(1.9), Inches(2.3), Inches(0.5))
            value_frame = value_box.text_frame
            value_frame.text = value
            value_frame.paragraphs[0].font.size = Pt(32)
            value_frame.paragraphs[0].font.bold = True
            value_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)  # Black
            value_frame.paragraphs[0].font.name = "Arial"
            value_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

        # Create "Events per Day" trend chart using matplotlib with filtered data
        chart_path = os.path.join(temp_dir, 'events_per_day_slide1.png')
        # Extract date range from the filtered data
        data_start_date = df_selected['Date'].min().date()
        data_end_date = df_selected['Date'].max().date()
        create_events_per_day_chart_for_slide1(df_selected, chart_path, data_start_date, data_end_date)

        # Add chart title
        chart_title = slide1.shapes.add_textbox(Inches(1.0), Inches(2.8), Inches(11), Inches(0.5))
        chart_title_frame = chart_title.text_frame
        chart_title_frame.text = "Events per Day"
        chart_title_frame.paragraphs[0].font.size = Pt(18)
        chart_title_frame.paragraphs[0].font.bold = True
        chart_title_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)
        chart_title_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        chart_title_frame.paragraphs[0].font.name = "Arial"

        # Add the Events per Day trend chart
        slide1.shapes.add_picture(chart_path, Inches(1.0), Inches(3.3), Inches(11.0), Inches(3.8))

        # SLIDE 2: Excel Data Comparison Analysis
        slide2 = prs.slides.add_slide(blank_slide_layout)

        # Set slide background to white
        background2 = slide2.background
        fill2 = background2.fill
        fill2.solid()
        fill2.fore_color.rgb = RGBColor(255, 255, 255)

        # Add main title centered at top
        main_title2 = slide2.shapes.add_textbox(Inches(1.0), Inches(0.3), Inches(11.33), Inches(0.8))
        main_title2_frame = main_title2.text_frame
        main_title2_frame.text = "📊Comparison Analysis"
        main_title2_frame.paragraphs[0].font.size = Pt(32)
        main_title2_frame.paragraphs[0].font.bold = True
        main_title2_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark gray
        main_title2_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        main_title2_frame.paragraphs[0].font.name = "Arial"

        # Add logo
        add_logo_to_slide(slide2, custom_logo_path)

        # Create comparison chart using actual data if available, otherwise use placeholder
        comparison_chart_path = os.path.join(temp_dir, 'comparison_chart_slide2.png')

        # Default values for metrics
        sample_avg1 = 11.86
        sample_avg2 = 9.29
        sample_percentage = -21.67
        chart_labels = ['Dataset 1 (Sample)', 'Dataset 2 (Sample)']

        if comparison_data and len(comparison_data) >= 5:
            # Use actual comparison data: [df1, df2, label1, label2, metrics]
            df1, df2, label1, label2, metrics = comparison_data
            chart_labels = [label1, label2]

            if metrics and len(metrics) >= 3:
                sample_avg1, sample_avg2, sample_percentage = metrics[:3]

            # Create chart with actual data
            try:
                chart_path, avg1, avg2, percentage_change = create_comparison_chart_for_ppt(
                    df1, df2, comparison_chart_path, label1, label2)
                if avg1 is not None and avg2 is not None:
                    sample_avg1, sample_avg2, sample_percentage = avg1, avg2, percentage_change
            except Exception as e:
                print(f"Error creating comparison chart with actual data: {e}")
                # Fall back to placeholder chart below

        if not comparison_data or not os.path.exists(comparison_chart_path):
            # Create placeholder chart for demonstration
            fig, ax = plt.subplots(figsize=(12, 5))

            # Create sample data for demonstration
            sample_dates = pd.date_range(start='2024-01-01', end='2024-01-07', freq='D')
            sample_data1 = [12, 8, 15, 10, 18, 6, 14]
            sample_data2 = [10, 6, 12, 8, 14, 4, 11]

            x_positions = range(len(sample_dates))
            date_labels = [d.strftime('%Y-%m-%d') for d in sample_dates]

            # Plot sample comparison lines
            ax.plot(x_positions, sample_data1, linewidth=3, color='#FF6B6B', alpha=0.8, label=chart_labels[0])
            ax.plot(x_positions, sample_data2, linewidth=3, color='#4ECDC4', alpha=0.8, label=chart_labels[1])

            # Add scatter points
            ax.scatter(x_positions, sample_data1, color='#FF6B6B', s=80, zorder=5, alpha=0.9)
            ax.scatter(x_positions, sample_data2, color='#4ECDC4', s=80, zorder=5, alpha=0.9)

            # Style the chart
            ax.set_xticks(x_positions)
            ax.set_xticklabels(date_labels, rotation=45, ha='right', fontsize=10)
            ax.set_title('📈 Average Fluctuation in Events',
                         fontsize=16, fontweight='bold', color='#1f2937', pad=20)
            ax.set_xlabel('Date', fontsize=12, fontweight='bold')
            ax.set_ylabel('Number of Events', fontsize=12, fontweight='bold')
            ax.legend(loc='upper right', frameon=True, fancybox=True, shadow=True, fontsize=11)
            ax.grid(True, alpha=0.3, linestyle='--')
            ax.set_facecolor('#fafafa')
            fig.patch.set_facecolor('white')

            plt.tight_layout()
            plt.savefig(comparison_chart_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()

        # Add the comparison chart
        slide2.shapes.add_picture(comparison_chart_path, Inches(0.8), Inches(1.3), Inches(11.7), Inches(3.8))

        # Add "📊 Comparison Metrics" section
        metrics_title = slide2.shapes.add_textbox(Inches(0.5), Inches(5.3), Inches(12), Inches(0.5))
        metrics_title_frame = metrics_title.text_frame
        metrics_title_frame.text = "📊 Comparison Metrics"
        metrics_title_frame.paragraphs[0].font.size = Pt(20)
        metrics_title_frame.paragraphs[0].font.bold = True
        metrics_title_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)
        metrics_title_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        metrics_title_frame.paragraphs[0].font.name = "Arial"

        # Add metrics cards using actual or sample data
        metrics_data = [
            (f"{chart_labels[0].split('(')[0].strip()} Average",
     f"{sample_avg1:.2f} events/day",
     RGBColor(
        254,
        242,
        242)),
            (f"{chart_labels[1].split('(')[0].strip()} Average",
         f"{sample_avg2:.2f} events/day",
         RGBColor(
            236,
            253,
            245)),
            ("Percentage Change", 
             f"{sample_percentage:+.1f}%",
             RGBColor(
                254,
                242,
                242) if sample_percentage < 0 else RGBColor(
                    236,
                    253,
                245)) ]

        for i, (label, value, bg_color) in enumerate(metrics_data):
            x_pos = 0.5 + (i * 4.0)

            # Create metric card
            metric_card = slide2.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x_pos),
                Inches(5.9),
                Inches(3.8),
                Inches(1.0))
            metric_card.fill.solid()
            metric_card.fill.fore_color.rgb = bg_color
            metric_card.line.color.rgb = RGBColor(229, 231, 235)
            metric_card.line.width = Pt(1)

            # Add label
            label_box = slide2.shapes.add_textbox(Inches(x_pos + 0.2), Inches(6.0), Inches(3.4), Inches(0.3))
            label_frame = label_box.text_frame
            label_frame.text = label
            label_frame.paragraphs[0].font.size = Pt(12)
            label_frame.paragraphs[0].font.bold = True
            label_frame.paragraphs[0].font.color.rgb = RGBColor(75, 85, 99)
            label_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add value
            value_box = slide2.shapes.add_textbox(Inches(x_pos + 0.2), Inches(6.4), Inches(3.4), Inches(0.4))
            value_frame = value_box.text_frame
            value_frame.text = value
            value_frame.paragraphs[0].font.size = Pt(16)
            value_frame.paragraphs[0].font.bold = True
            value_frame.paragraphs[0].font.color.rgb = RGBColor(17, 24, 39)
            value_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

        # SLIDE 3: High Impact Analysis (matching the screenshot exactly)
        slide3 = prs.slides.add_slide(blank_slide_layout)

        # Set slide background to white
        background3 = slide3.background
        fill3 = background3.fill
        fill3.solid()
        fill3.fore_color.rgb = RGBColor(255, 255, 255)

        # Add main title "High Impact Analysis" in top left
        main_title3 = slide3.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(6), Inches(0.8))
        main_title3_frame = main_title3.text_frame
        main_title3_frame.text = "High Impact Analysis"
        main_title3_frame.paragraphs[0].font.size = Pt(36)
        main_title3_frame.paragraphs[0].font.bold = True
        main_title3_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark gray
        main_title3_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        main_title3_frame.paragraphs[0].font.name = "Arial"

        # Add logo (VISIONIFY by default or custom if provided)
        add_logo_to_slide(slide3, custom_logo_path)

        # Use accurate severity data (already calculated above)
        # Create 3 impact metric cards using accurate centralized data
        impact_cards = [
            ("Near-Miss Events",
             severity_data['near_miss'],
             f"{severity_data['near_miss_percentage']:.1f}% of total events",
             RGBColor(220,
             252,
             231),
             RGBColor(34,
             197,
             94),
             1.0),
            # Green
            ("Emergency Events", severity_data['emergency'], f"{severity_data['emergency_percentage']:.1f}% of total events", RGBColor(254, 242, 242), RGBColor(239, 68, 68), 5.0),  # Red
            ("High Risk Events", severity_data['high_risk'], "Combined critical incidents", RGBColor(254, 242, 242), RGBColor(239, 68, 68), 9.0)  # Red
        ]

        for title, value, subtitle, bg_color, text_color, x_pos in impact_cards:
            # Create impact card with colored background
            impact_card = slide3.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x_pos),
                Inches(1.3),
                Inches(3.5),
                Inches(1.2))
            impact_card.fill.solid()
            impact_card.fill.fore_color.rgb = bg_color
            impact_card.line.color.rgb = RGBColor(229, 231, 235)  # Light border
            impact_card.line.width = Pt(1)

            # Add title text (bold, dark)
            title_box = slide3.shapes.add_textbox(Inches(x_pos + 0.2), Inches(1.5), Inches(3.1), Inches(0.4))
            title_frame = title_box.text_frame
            title_frame.text = title
            title_frame.paragraphs[0].font.size = Pt(14)
            title_frame.paragraphs[0].font.bold = True
            title_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark gray
            title_frame.paragraphs[0].font.name = "Arial"
            title_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add value text (large, bold, colored)
            value_box = slide3.shapes.add_textbox(Inches(x_pos + 0.2), Inches(1.8), Inches(3.1), Inches(0.4))
            value_frame = value_box.text_frame
            value_frame.text = str(value)
            value_frame.paragraphs[0].font.size = Pt(28)
            value_frame.paragraphs[0].font.bold = True
            value_frame.paragraphs[0].font.color.rgb = text_color
            value_frame.paragraphs[0].font.name = "Arial"
            value_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add subtitle text (smaller, gray)
            subtitle_box = slide3.shapes.add_textbox(Inches(x_pos + 0.2), Inches(2.2), Inches(3.1), Inches(0.3))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = subtitle
            subtitle_frame.paragraphs[0].font.size = Pt(10)
            subtitle_frame.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)  # Gray text
            subtitle_frame.paragraphs[0].font.name = "Arial"
            subtitle_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

        # Add "Top 3 Contributing Cameras" title (moved to left side)
        cameras_title = slide3.shapes.add_textbox(Inches(0.5), Inches(2.8), Inches(6), Inches(0.5))
        cameras_title_frame = cameras_title.text_frame
        cameras_title_frame.text = "Top 3 Contributing Cameras"
        cameras_title_frame.paragraphs[0].font.size = Pt(20)
        cameras_title_frame.paragraphs[0].font.bold = True
        cameras_title_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)
        cameras_title_frame.paragraphs[0].alignment = PP_ALIGN.LEFT  # Changed from CENTER to LEFT
        cameras_title_frame.paragraphs[0].font.name = "Arial"

        # Use camera_contributions data (same as dashboard)
        # Ensure we have exactly 3 camera entries for display
        display_cameras = camera_severity_data.copy()
        while len(display_cameras) < 3:
            display_cameras.append({
                'camera': f'Camera-{len(display_cameras)+1}',
                'location': 'No Data Available',
                'total': 0,
                'high_severity': 0
            })

        # Create 3 camera cards using dashboard format data
        for i, camera_info in enumerate(display_cameras[:3]):
            x_position = 1.0 + (i * 4.0)  # Space cards evenly

            # Create camera card
            camera_card = slide3.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(x_position),
                Inches(3.4),
                Inches(3.8),
                Inches(2.2))
            camera_card.fill.solid()
            camera_card.fill.fore_color.rgb = RGBColor(248, 250, 252)  # Very light gray
            camera_card.line.color.rgb = RGBColor(229, 231, 235)  # Light border
            camera_card.line.width = Pt(1)

            # Add rank badge (circular blue)
            rank_circle = slide3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(
                x_position + 0.2), Inches(3.6), Inches(0.6), Inches(0.6))
            rank_circle.fill.solid()
            rank_circle.fill.fore_color.rgb = RGBColor(59, 130, 246)  # Blue
            rank_circle.line.fill.background()

            # Add rank number
            rank_text = slide3.shapes.add_textbox(Inches(x_position + 0.35), Inches(3.75), Inches(0.3), Inches(0.3))
            rank_text_frame = rank_text.text_frame
            rank_text_frame.text = f"#{i+1}"
            rank_text_frame.paragraphs[0].font.size = Pt(14)
            rank_text_frame.paragraphs[0].font.bold = True
            rank_text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)  # White
            rank_text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add high severity badge (top right)
            severity_badge = slide3.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(
        x_position + 2.8), Inches(3.6), Inches(0.9), Inches(0.3))
            severity_badge.fill.solid()
            severity_badge.fill.fore_color.rgb = RGBColor(219, 234, 254)  # Light blue
            severity_badge.line.color.rgb = RGBColor(59, 130, 246)  # Blue border
            severity_badge.line.width = Pt(1)

            # Add severity text using dashboard format data
            severity_text = slide3.shapes.add_textbox(Inches(x_position + 2.85), Inches(3.65), Inches(0.8), Inches(0.2))
            severity_text_frame = severity_text.text_frame
            severity_text_frame.text = f"{camera_info['high_severity']} High Severity"
            severity_text_frame.paragraphs[0].font.size = Pt(8)
            severity_text_frame.paragraphs[0].font.color.rgb = RGBColor(59, 130, 246)  # Blue
            severity_text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add camera name (large, bold) using dashboard format data
            camera_name = slide3.shapes.add_textbox(Inches(x_position + 0.2), Inches(4.3), Inches(3.4), Inches(0.4))
            camera_name_frame = camera_name.text_frame
            camera_name_frame.text = camera_info['camera']
            camera_name_frame.paragraphs[0].font.size = Pt(16)
            camera_name_frame.paragraphs[0].font.bold = True
            camera_name_frame.paragraphs[0].font.color.rgb = RGBColor(59, 130, 246)  # Blue
            camera_name_frame.paragraphs[0].font.name = "Arial"
            camera_name_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add area/zone (smaller, gray) using dashboard format data
            area_text = slide3.shapes.add_textbox(Inches(x_position + 0.2), Inches(4.7), Inches(3.4), Inches(0.3))
            area_text_frame = area_text.text_frame
            area_text_frame.text = camera_info['location']
            area_text_frame.paragraphs[0].font.size = Pt(11)
            area_text_frame.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)  # Gray
            area_text_frame.paragraphs[0].font.name = "Arial"
            area_text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Add event count (large, blue) using dashboard format data
            event_count = slide3.shapes.add_textbox(Inches(x_position + 0.2), Inches(5.0), Inches(3.4), Inches(0.4))
            event_count_frame = event_count.text_frame
            event_count_frame.text = f"{camera_info['total']} events"
            event_count_frame.paragraphs[0].font.size = Pt(18)
            event_count_frame.paragraphs[0].font.bold = True
            event_count_frame.paragraphs[0].font.color.rgb = RGBColor(59, 130, 246)  # Blue
            event_count_frame.paragraphs[0].font.name = "Arial"
            event_count_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

        # SLIDE 4: Top 10 Cameras By Event Count Table (using dashboard logic)
        slide4 = prs.slides.add_slide(blank_slide_layout)

        # Set slide background to white
        background4 = slide4.background
        fill4 = background4.fill
        fill4.solid()
        fill4.fore_color.rgb = RGBColor(255, 255, 255)

        # Add main title "Top 10 Cameras By Event Count" in top left
        main_title4 = slide4.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.8))
        main_title4_frame = main_title4.text_frame
        main_title4_frame.text = "Top 10 Cameras By Event Count"
        main_title4_frame.paragraphs[0].font.size = Pt(36)
        main_title4_frame.paragraphs[0].font.bold = True
        main_title4_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark gray
        main_title4_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        main_title4_frame.paragraphs[0].font.name = "Arial"

        # Add logo (VISIONIFY by default or custom if provided)
        add_logo_to_slide(slide4, custom_logo_path)

        # Use the same logic as dashboard to create the heatmap table
        camera_column = None
        if 'Camera' in df_selected.columns:
            camera_column = 'Camera'
        elif 'Area' in df_selected.columns:
            camera_column = 'Area'

        if camera_column and len(df_selected) > 0 and 'Events' in df_selected.columns:
            # Get top 10 cameras by total event count
            top_cameras = df_selected[camera_column].value_counts().head(10)
            top_camera_names = top_cameras.index.tolist()

            # Get top 5 event types by total count
            top_events_all = df_selected['Events'].value_counts().head(5)
            top_event_names_all = top_events_all.index.tolist()

            # Create pivot table from ALL data first
            full_camera_pivot = df_selected.groupby([camera_column, 'Events']).size().unstack(fill_value=0)

            if not full_camera_pivot.empty:
                # Filter to only show top 10 cameras as rows
                camera_event_pivot = full_camera_pivot.loc[full_camera_pivot.index.intersection(top_camera_names)]

                # Ensure we have all top 10 cameras as rows (add missing ones with zeros)
                for camera in top_camera_names:
                    if camera not in camera_event_pivot.index:
                        camera_event_pivot.loc[camera] = 0

                # Reorder rows by total count (descending)
                camera_event_pivot['Total'] = camera_event_pivot.sum(axis=1)
                camera_event_pivot = camera_event_pivot.sort_values('Total', ascending=False).drop('Total', axis=1)

                # Ensure we have all top 5 event types as columns (add missing ones with zeros)
                for event in top_event_names_all:
                    if event not in camera_event_pivot.columns:
                        camera_event_pivot[event] = 0

                # Keep only top 5 event types and reorder by their total counts
                available_top_events = [evt for evt in top_event_names_all if evt in camera_event_pivot.columns]
                camera_event_pivot = camera_event_pivot[available_top_events]

                if not camera_event_pivot.empty:
                    # Create table with the same data as dashboard
                    table_rows = len(camera_event_pivot) + 1  # +1 for header
                    table_cols = len(camera_event_pivot.columns) + 1  # +1 for camera name column

                    # Create table (positioned below title with proper margins)
                    table_left = Inches(0.8)
                    table_top = Inches(1.8)
                    table_width = Inches(11.7)  # Reduced width to fit within slide bounds
                    table_height = Inches(5.5)  # Reduced height to ensure fit

                    table = slide4.shapes.add_table(
                        table_rows,
                        table_cols,
                        table_left,
                        table_top,
                        table_width,
                        table_height).table

                    # Set table style with enhanced borders to match dashboard
                    table.style = 'TableGrid'  # Basic grid style as foundation
                    
                    # Apply custom border styling to all cells for neat appearance
                    from pptx.enum.dml import MSO_LINE
                    from pptx.enum.shapes import MSO_CONNECTOR
                    
                    # Calculate max and min values for color scaling (same as dashboard)
                    df_for_scaling = camera_event_pivot.copy()
                    vmax = df_for_scaling.max().max() if not df_for_scaling.empty else 1
                    vmin = df_for_scaling.min().min() if not df_for_scaling.empty else 0

                    # Header row - match dashboard styling exactly
                    header_cells = table.rows[0].cells
                    header_cells[0].text = "Camera/Area"
                    # Header styling - white background, dark text, bold
                    header_cells[0].fill.solid()
                    header_cells[0].fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
                    header_cells[0].text_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark text
                    header_cells[0].text_frame.paragraphs[0].font.bold = True
                    header_cells[0].text_frame.paragraphs[0].font.size = Pt(12)
                    header_cells[0].text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
                    header_cells[0].text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

                    # Event type headers - match dashboard exactly
                    for j, event_type in enumerate(camera_event_pivot.columns):
                        header_cells[j + 1].text = event_type
                        # Header styling - white background, dark text, bold, center aligned
                        header_cells[j + 1].fill.solid()
                        header_cells[j + 1].fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
                        text_frame = header_cells[j + 1].text_frame
                        text_frame.word_wrap = True
                        text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                        text_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark text
                        text_frame.paragraphs[0].font.bold = True
                        text_frame.paragraphs[0].font.size = Pt(12)
                        text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

                    # Data rows - match dashboard exactly
                    for i, camera_name in enumerate(camera_event_pivot.index):
                        row_cells = table.rows[i + 1].cells

                        # Camera name cell - white background, dark text, bold, left aligned
                        row_cells[0].text = camera_name
                        row_cells[0].fill.solid()
                        row_cells[0].fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
                        camera_text_frame = row_cells[0].text_frame
                        camera_text_frame.word_wrap = True
                        camera_text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                        camera_text_frame.paragraphs[0].font.bold = True
                        camera_text_frame.paragraphs[0].font.size = Pt(12)
                        camera_text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
                        camera_text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                        camera_text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)  # Black text

                        # Event count cells with heatmap colors - match dashboard exactly
                        for j, event_type in enumerate(camera_event_pivot.columns):
                            value = camera_event_pivot.loc[camera_name, event_type]
                            row_cells[j + 1].text = str(value)

                            # Apply heatmap color using exact same logic as dashboard
                            if vmax > vmin and value > 0:
                                intensity = (value - vmin) / (vmax - vmin)
                            else:
                                intensity = 0

                            # Color scale from white (0) to red (max) - exact same as dashboard
                            if value == 0:
                                bg_color = RGBColor(255, 255, 255)  # White
                                text_color = RGBColor(108, 117, 125)  # Gray
                            else:
                                # Red intensity based on value - exact same as dashboard
                                red = 255
                                green = int(255 - (intensity * 153))  # 255 to 102
                                blue = int(255 - (intensity * 153))   # 255 to 102
                                bg_color = RGBColor(red, green, blue)
                                text_color = RGBColor(0, 0, 0) if intensity < 0.7 else RGBColor(255, 255, 255)
                            
                            row_cells[j + 1].fill.solid()
                            row_cells[j + 1].fill.fore_color.rgb = bg_color
                            row_cells[j + 1].text_frame.paragraphs[0].font.size = Pt(12)
                            row_cells[j + 1].text_frame.paragraphs[0].font.bold = True
                            row_cells[j + 1].text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                            row_cells[j + 1].text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                            row_cells[j + 1].text_frame.paragraphs[0].font.color.rgb = text_color

                    # Apply enhanced border styling to all cells for neat appearance
                    try:
                        for row in table.rows:
                            for cell in row.cells:
                                # Set border properties for each cell with error handling
                                try:
                                    # Top border - thicker and darker
                                    if hasattr(cell, 'border_top') and hasattr(cell.border_top, 'color'):
                                        cell.border_top.color.rgb = RGBColor(156, 163, 175)  # Darker gray border
                                    if hasattr(cell, 'border_top') and hasattr(cell.border_top, 'width'):
                                        cell.border_top.width = Pt(3)  # Thicker border
                                    
                                    # Bottom border - thicker and darker
                                    if hasattr(cell, 'border_bottom') and hasattr(cell.border_bottom, 'color'):
                                        cell.border_bottom.color.rgb = RGBColor(156, 163, 175)
                                    if hasattr(cell, 'border_bottom') and hasattr(cell.border_bottom, 'width'):
                                        cell.border_bottom.width = Pt(3)  # Thicker border
                                    
                                    # Left border - thicker and darker
                                    if hasattr(cell, 'border_left') and hasattr(cell.border_left, 'color'):
                                        cell.border_left.color.rgb = RGBColor(156, 163, 175)
                                    if hasattr(cell, 'border_left') and hasattr(cell.border_left, 'width'):
                                        cell.border_left.width = Pt(3)  # Thicker border
                                    
                                    # Right border - thicker and darker
                                    if hasattr(cell, 'border_right') and hasattr(cell.border_right, 'color'):
                                        cell.border_right.color.rgb = RGBColor(156, 163, 175)
                                    if hasattr(cell, 'border_right') and hasattr(cell.border_right, 'width'):
                                        cell.border_right.width = Pt(3)  # Thicker border
                                except Exception as e:
                                    # Skip border styling if not supported
                                    continue
                    except Exception as e:
                        # Skip border styling entirely if table doesn't support it
                        pass
                else:
                    # No data available message
                    no_data_text = slide4.shapes.add_textbox(Inches(2), Inches(3), Inches(8), Inches(2))
                    no_data_frame = no_data_text.text_frame
                    no_data_frame.text = "No event data available for cameras in the selected period"
                    no_data_frame.paragraphs[0].font.size = Pt(18)
                    no_data_frame.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)
                    no_data_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            else:
                # No data available message
                no_data_text = slide4.shapes.add_textbox(Inches(2), Inches(3), Inches(8), Inches(2))
                no_data_frame = no_data_text.text_frame
                no_data_frame.text = "No camera event data available for the selected period"
                no_data_frame.paragraphs[0].font.size = Pt(18)
                no_data_frame.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)
                no_data_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        else:
            # No data available message
            no_data_text = slide4.shapes.add_textbox(Inches(2), Inches(3), Inches(8), Inches(2))
            no_data_frame = no_data_text.text_frame
            no_data_frame.text = "No camera/area data available for the selected period"
            no_data_frame.paragraphs[0].font.size = Pt(18)
            no_data_frame.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)
            no_data_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

        # SLIDE 5: Events by Type and Events by Area Charts (side by side)
        slide5 = prs.slides.add_slide(blank_slide_layout)

        # Set slide background to white
        background5 = slide5.background
        fill5 = background5.fill
        fill5.solid()
        fill5.fore_color.rgb = RGBColor(255, 255, 255)

        # Add main title "Events Analysis" in top left
        main_title5 = slide5.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.8))
        main_title5_frame = main_title5.text_frame
        main_title5_frame.text = "Events Analysis"
        main_title5_frame.paragraphs[0].font.size = Pt(36)
        main_title5_frame.paragraphs[0].font.bold = True
        main_title5_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark gray
        main_title5_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        main_title5_frame.paragraphs[0].font.name = "Arial"

        # Add logo (VISIONIFY by default or custom if provided)
        add_logo_to_slide(slide5, custom_logo_path)

        # Create Events by Type chart (left side)
        events_by_type_chart_path = os.path.join(temp_dir, 'events_by_type_slide5.png')
        create_events_by_type_chart_for_slide5(df_selected, events_by_type_chart_path, event_data)

        # Create Events by Area chart (right side) using filtered data
        events_by_area_chart_path = os.path.join(temp_dir, 'events_by_area_slide5.png')
        create_events_by_area_chart_for_slide5(df_selected, events_by_area_chart_path, camera_data)

        # Note: Chart titles removed as requested - no individual chart headers

        # Add the Events by Type bar chart (left side)
        slide5.shapes.add_picture(events_by_type_chart_path, Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.5))

        # Add the Events by Area bar chart (right side)
        slide5.shapes.add_picture(events_by_area_chart_path, Inches(7.0), Inches(1.8), Inches(5.5), Inches(4.5))

        # SLIDE 6: Recommendations (Clean slide for manual recommendations entry)
        slide6 = prs.slides.add_slide(blank_slide_layout)

        # Set slide background to white
        background6 = slide6.background
        fill6 = background6.fill
        fill6.solid()
        fill6.fore_color.rgb = RGBColor(255, 255, 255)

        # Add main title "Recommendations" in top left
        main_title6 = slide6.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.8))
        main_title6_frame = main_title6.text_frame
        main_title6_frame.text = "Recommendations"
        main_title6_frame.paragraphs[0].font.size = Pt(36)
        main_title6_frame.paragraphs[0].font.bold = True
        main_title6_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)  # Dark gray
        main_title6_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        main_title6_frame.paragraphs[0].font.name = "Arial"

        # Add logo (VISIONIFY by default or custom if provided)
        add_logo_to_slide(slide6, custom_logo_path)

        # Add styled content areas for manual recommendations entry
        # Create three professional recommendation boxes
        
        # Recommendation Box 1 - High Priority
        rec_box1 = slide6.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(1.5), Inches(12.3), Inches(1.3)
        )
        rec_box1.fill.solid()
        rec_box1.fill.fore_color.rgb = RGBColor(254, 242, 242)  # Light red background
        rec_box1.line.color.rgb = RGBColor(239, 68, 68)  # Red border
        rec_box1.line.width = Pt(2)
        
        # Add high priority recommendation text
        rec_text1 = slide6.shapes.add_textbox(Inches(0.7), Inches(1.6), Inches(11.9), Inches(1.1))
        rec_frame1 = rec_text1.text_frame
        rec_frame1.word_wrap = True
        rec_frame1.margin_left = Inches(0.1)
        rec_frame1.margin_top = Inches(0.1)
        
        p1 = rec_frame1.paragraphs[0]
        p1.text = "🔴 HIGH PRIORITY RECOMMENDATION"
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(239, 68, 68)  # Red text
        p1.font.name = "Arial"
        
        p2 = rec_frame1.add_paragraph()
        p2.text = "• Click here to add your high priority safety recommendations"
        p2.font.size = Pt(12)
        p2.font.color.rgb = RGBColor(107, 114, 128)  # Gray text
        p2.font.name = "Arial"
        p2.level = 1
        
        # Recommendation Box 2 - Medium Priority
        rec_box2 = slide6.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(3.0), Inches(12.3), Inches(1.3)
        )
        rec_box2.fill.solid()
        rec_box2.fill.fore_color.rgb = RGBColor(255, 251, 235)  # Light yellow background
        rec_box2.line.color.rgb = RGBColor(245, 158, 11)  # Yellow border
        rec_box2.line.width = Pt(2)
        
        # Add medium priority recommendation text
        rec_text2 = slide6.shapes.add_textbox(Inches(0.7), Inches(3.1), Inches(11.9), Inches(1.1))
        rec_frame2 = rec_text2.text_frame
        rec_frame2.word_wrap = True
        rec_frame2.margin_left = Inches(0.1)
        rec_frame2.margin_top = Inches(0.1)
        
        p3 = rec_frame2.paragraphs[0]
        p3.text = "🟡 MEDIUM PRIORITY RECOMMENDATION"
        p3.font.size = Pt(14)
        p3.font.bold = True
        p3.font.color.rgb = RGBColor(245, 158, 11)  # Yellow text
        p3.font.name = "Arial"
        
        p4 = rec_frame2.add_paragraph()
        p4.text = "• Click here to add your medium priority improvement suggestions"
        p4.font.size = Pt(12)
        p4.font.color.rgb = RGBColor(107, 114, 128)  # Gray text
        p4.font.name = "Arial"
        p4.level = 1
        
        # Recommendation Box 3 - General Improvements
        rec_box3 = slide6.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(4.5), Inches(12.3), Inches(1.3)
        )
        rec_box3.fill.solid()
        rec_box3.fill.fore_color.rgb = RGBColor(240, 253, 244)  # Light green background
        rec_box3.line.color.rgb = RGBColor(34, 197, 94)  # Green border
        rec_box3.line.width = Pt(2)
        
        # Add general improvement recommendation text
        rec_text3 = slide6.shapes.add_textbox(Inches(0.7), Inches(4.6), Inches(11.9), Inches(1.1))
        rec_frame3 = rec_text3.text_frame
        rec_frame3.word_wrap = True
        rec_frame3.margin_left = Inches(0.1)
        rec_frame3.margin_top = Inches(0.1)
        
        p5 = rec_frame3.paragraphs[0]
        p5.text = "🟢 GENERAL IMPROVEMENTS"
        p5.font.size = Pt(14)
        p5.font.bold = True
        p5.font.color.rgb = RGBColor(34, 197, 94)  # Green text
        p5.font.name = "Arial"
        
        p6 = rec_frame3.add_paragraph()
        p6.text = "• Click here to add general operational improvements and best practices"
        p6.font.size = Pt(12)
        p6.font.color.rgb = RGBColor(107, 114, 128)  # Gray text
        p6.font.name = "Arial"
        p6.level = 1
        
        # Add footer note
        footer_note = slide6.shapes.add_textbox(Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.8))
        footer_frame = footer_note.text_frame
        footer_frame.text = "💡 TIP: Double-click on each section above to edit and add your specific recommendations based on the analyzed data."
        footer_frame.paragraphs[0].font.size = Pt(11)
        footer_frame.paragraphs[0].font.italic = True
        footer_frame.paragraphs[0].font.color.rgb = RGBColor(107, 114, 128)  # Gray text
        footer_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        footer_frame.paragraphs[0].font.name = "Arial"
        
        # Add default slides at their original positions before saving
        add_default_slides(prs, blank_slide_layout)

        # Save to buffer
        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)

        return ppt_buffer

    finally:
        # Clean up temporary files
        shutil.rmtree(temp_dir, ignore_errors=True)

    if camera_column and not df.empty:
        # Get area/camera counts
        area_counts = df[camera_column].value_counts()

        if len(area_counts) > 0:
            # Get top 10 areas only
            top_area_counts = area_counts.head(10)

            # Create professional blue bars
            bars = ax.bar(range(len(top_area_counts)), top_area_counts.values,
                          color=PROFESSIONAL_BLUE, width=0.7, alpha=0.9)

            # Add value labels on top of each bar with percentages
            total_events = top_area_counts.sum()
            for i, (bar, count) in enumerate(zip(bars, top_area_counts.values)):
                percentage = (count / total_events * 100) if total_events > 0 else 0
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(top_area_counts.values) * 0.02,
                        f'{count}\n({percentage:.1f}%)', ha='center', va='bottom',
                        fontweight='bold', fontsize=10, color='#333333')

            # Customize the chart
            ax.set_xlabel('Area (Top 10)', fontsize=12, color='#333333', fontweight='bold')
            ax.set_ylabel('Event Count', fontsize=12, color='#333333', fontweight='bold')
            ax.set_title('Events by Area (Top 10)', fontsize=14, color='#333333', fontweight='bold', pad=20)

            # Set x-axis labels with proper formatting
            ax.set_xticks(range(len(top_area_counts)))
            labels = []
            for label in top_area_counts.index:
                # Format labels to be more readable
                if len(label) > 12:
                    # Split long labels
                    words = str(label).split()
                    if len(words) > 1:
                        mid = len(words) // 2
                        labels.append('\n'.join([' '.join(words[:mid]), ' '.join(words[mid:])]))
                    else:
                        labels.append(str(label)[:12] + '...')
                else:
                    labels.append(str(label))
            ax.set_xticklabels(labels, fontsize=9, color='#333333', rotation=45, ha='right')

            # Style improvements
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#E5E7EB')
            ax.spines['bottom'].set_color('#E5E7EB')

            # Set y-axis to start from 0 and add some padding
            ax.set_ylim(0, max(top_area_counts.values) * 1.15)

            # Style tick parameters
            ax.tick_params(axis='both', which='major', labelsize=9, colors='#333333')

            # Add grid for better readability
            ax.grid(True, alpha=0.3, axis='y')
            ax.set_axisbelow(True)
        else:
            # No areas found
            ax.text(0.5, 0.5, 'No Area data available', ha='center', va='center',
                    transform=ax.transAxes, fontsize=14)
    else:
        # No camera/area column
        ax.text(0.5, 0.5, 'No Area/Camera data available', ha='center', va='center',
                transform=ax.transAxes, fontsize=14)

    # Save with high DPI for PowerPoint
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def get_accurate_event_counts(df):
    """
    Get accurate event counts for all visualizations
    Returns standardized event data used by both dashboard and PowerPoint
    """
    if df.empty or 'Events' not in df.columns:
        return {}

    # Get exact event counts
    event_counts = df['Events'].value_counts()
    total_events = len(df)

    return {
        'event_counts': event_counts,
        'total_events': total_events,
        'event_types': len(event_counts),
        'top_5_events': event_counts.head(5),
        'all_events': event_counts
    }


def get_accurate_camera_data(df):
    """
    Get accurate camera/area data for all visualizations
    Returns standardized camera data used by both dashboard and PowerPoint
    """
    camera_data = {
        'camera_column': None,
        'camera_counts': {},
        'top_cameras': {},
        'total_cameras': 0
    }

    # Determine camera column (Camera > Area > None) - Prioritize individual cameras
    if 'Camera' in df.columns:
        camera_data['camera_column'] = 'Camera'
    elif 'Area' in df.columns:
        camera_data['camera_column'] = 'Area'
    else:
        return camera_data

    if camera_data['camera_column'] and not df.empty:
        camera_counts = df[camera_data['camera_column']].value_counts()
        camera_data['camera_counts'] = camera_counts
        camera_data['top_cameras'] = camera_counts.head(10)
        camera_data['total_cameras'] = len(camera_counts)
        camera_data['top_3_cameras'] = camera_counts.head(3)
        # Add areas/cameras list for area-wise reports
        camera_data['areas'] = camera_counts.index.tolist()
        camera_data['counts'] = camera_counts.values.tolist()

    return camera_data


def get_accurate_severity_data(df):
    """
    Get accurate severity analysis for all visualizations
    Returns standardized severity data used by both dashboard and PowerPoint
    """
    severity_data = {
        'has_severity_column': False,
        'near_miss': 0,
        'emergency': 0,
        'high_risk': 0,
        'near_miss_percentage': 0.0,
        'emergency_percentage': 0.0,
        'high_risk_percentage': 0.0,
        'severity_breakdown': {}
    }

    if df.empty:
        return severity_data

    total_events = len(df)

    # Check for Events column patterns first
    if 'Events' in df.columns:
        # Near-Miss Events (case-insensitive, flexible matching)
        near_miss_pattern = r'near[-\s]*miss'
        near_miss_mask = df['Events'].str.contains(near_miss_pattern, case=False, na=False, regex=True)
        severity_data['near_miss'] = near_miss_mask.sum()

        # Emergency Events (case-insensitive)
        emergency_pattern = r'emergency'
        emergency_mask = df['Events'].str.contains(emergency_pattern, case=False, na=False, regex=True)
        severity_data['emergency'] = emergency_mask.sum()

    # Check for Severity column
    if 'Severity' in df.columns:
        severity_data['has_severity_column'] = True
        severity_str = df['Severity'].astype(str).str.lower()

        # High Risk Events (critical, high, severe, major)
        high_risk_pattern = r'(critical|high|severe|major)'
        high_risk_mask = severity_str.str.contains(high_risk_pattern, case=False, na=False, regex=True)
        severity_data['high_risk'] = high_risk_mask.sum()

        # Get severity breakdown
        severity_counts = df['Severity'].value_counts()
        severity_data['severity_breakdown'] = severity_counts

    # Calculate percentages
    if total_events > 0:
        severity_data['near_miss_percentage'] = (severity_data['near_miss'] / total_events) * 100
        severity_data['emergency_percentage'] = (severity_data['emergency'] / total_events) * 100
        severity_data['high_risk_percentage'] = (severity_data['high_risk'] / total_events) * 100

    return severity_data


def get_accurate_kpis(df):
    """
    Get accurate KPIs for dashboard and PowerPoint
    Returns comprehensive KPI data ensuring consistency
    """
    event_data = get_accurate_event_counts(df)
    camera_data = get_accurate_camera_data(df)
    severity_data = get_accurate_severity_data(df)

    # Calculate date range
    date_range = "No date data"
    if 'Date' in df.columns and not df.empty:
        try:
            if not pd.api.types.is_datetime64_any_dtype(df['Date']):
                df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')

            valid_dates = df['Date'].dropna()
            if not valid_dates.empty:
                start_date = valid_dates.min().strftime('%B %d, %Y')
                end_date = valid_dates.max().strftime('%B %d, %Y')
                if start_date == end_date:
                    date_range = start_date
                else:
                    date_range = f"{start_date} to {end_date}"
        except BaseException:
            date_range = "Date parsing error"

    # Calculate areas monitored using EXACT same logic as dashboard
    if df.empty:
        total_areas = 0
    elif 'Area' in df.columns:
        total_areas = df['Area'].nunique()
    elif 'Camera' in df.columns:
        total_areas = df['Camera'].nunique()
    else:
        total_areas = 0

    # Calculate Review Percent and Accuracy Percent
    # Review Percent: Percentage of events that have been reviewed based on "Reviewed" column with Yes/No values
    review_percent = 85  # Default value - should be calculated from actual review data
    if 'Reviewed' in df.columns:
        # Count "Yes" values in the Reviewed column
        reviewed_yes_count = df['Reviewed'].str.contains('yes', case=False, na=False).sum()
        review_percent = round((reviewed_yes_count / len(df)) * 100) if len(df) > 0 else 0
    elif 'Review' in df.columns:
        reviewed_events = df['Review'].notna().sum()
        review_percent = round((reviewed_events / len(df)) * 100) if len(df) > 0 else 0
    elif 'Status' in df.columns:
        # Alternative: use Status column to determine reviewed events
        reviewed_events = df['Status'].str.contains('review|complete|resolved', case=False, na=False).sum()
        review_percent = round((reviewed_events / len(df)) * 100) if len(df) > 0 else 0

    return {
        'total_events': event_data.get('total_events', 0),
        'total_event_types': event_data.get('event_types', 0),
        'total_cameras': camera_data.get('total_cameras', 0),
        'total_areas': total_areas,  # Use dashboard calculation logic
        'review_percent': review_percent,
        'date_range': date_range,
        'near_miss_events': severity_data.get('near_miss', 0),
        'emergency_events': severity_data.get('emergency', 0),
        'high_risk_events': severity_data.get('high_risk', 0),
        'near_miss_percentage': severity_data.get('near_miss_percentage', 0.0),
        'emergency_percentage': severity_data.get('emergency_percentage', 0.0),
        'has_severity_data': severity_data.get('has_severity_column', False)
    }

# Updated PowerPoint generation functions using accurate data


def create_camera_event_heatmap_table(df, event_data, camera_data):
    """
    Create camera event heatmap table data for PowerPoint slide
    This function replicates the EXACT dashboard table creation logic
    """
    # Initialize empty return structure
    table_info = {
        'camera_names': [],
        'event_types': [],
        'heatmap_matrix': [],
        'max_count': 0
    }

    if df.empty or not camera_data['camera_column'] or 'Events' not in df.columns:
        return table_info

    camera_column = camera_data['camera_column']

    try:
        # STEP 1: EXACTLY replicate dashboard logic - Get top 10 cameras by total event count
        print(f"\n=== DETAILED DEBUG: Camera Analysis ===")
        print(f"Total rows in df: {len(df)}")
        print(f"Camera column name: '{camera_column}'")
        print(f"Unique cameras in dataset: {df[camera_column].nunique()}")

        # Show ALL unique camera names from the dataset
        all_cameras = df[camera_column].value_counts()
        print(f"\nALL cameras in your data (showing first 20):")
        for i, (camera, count) in enumerate(all_cameras.head(20).items()):
            print(f"  {i+1}. '{camera}': {count} events")

        top_cameras = df[camera_column].value_counts().head(10)
        top_camera_names = top_cameras.index.tolist()

        print(f"\nTop 10 cameras selected for PowerPoint:")
        for i, (camera, count) in enumerate(top_cameras.items()):
            print(f"  {i+1}. '{camera}': {count} events")

        # STEP 2: Get top 5 event types by total count from all data
        top_events_all = df['Events'].value_counts().head(5)
        top_event_names_all = top_events_all.index.tolist()

        if not top_camera_names or not top_event_names_all:
            return table_info

        # STEP 3: EXACTLY replicate dashboard pivot table creation
        # Create pivot table from ALL data first (not filtering by cameras yet)
        full_camera_pivot = df.groupby([camera_column, 'Events']).size().unstack(fill_value=0)

        if not full_camera_pivot.empty:
            # STEP 4: Filter to only show top 10 cameras as rows
            camera_event_pivot = full_camera_pivot.loc[full_camera_pivot.index.intersection(
                top_camera_names)]

            # STEP 5: Ensure we have all top 10 cameras as rows (add missing
            # ones with zeros)
            for camera in top_camera_names:
                if camera not in camera_event_pivot.index:
                    # Add row of zeros for this camera (EXACT dashboard method)
                    camera_event_pivot.loc[camera] = 0

            # STEP 6: Reorder rows by total count (descending) - EXACT dashboard method
            camera_event_pivot['Total'] = camera_event_pivot.sum(axis=1)
            camera_event_pivot = camera_event_pivot.sort_values('Total', ascending=False).drop('Total', axis=1)

            # STEP 7: Ensure we have all top 5 event types as columns (add missing ones with zeros)
            for event in top_event_names_all:
                if event not in camera_event_pivot.columns:
                    camera_event_pivot[event] = 0

            # STEP 8: Keep only top 5 event types and reorder by their total counts
            available_top_events = [evt for evt in top_event_names_all if evt in camera_event_pivot.columns]
            camera_event_pivot = camera_event_pivot[available_top_events]

            # STEP 9: Convert to PowerPoint format - FORCE individual camera names (no grouping allowed)
            individual_camera_names = camera_event_pivot.index.tolist()

            # DEBUG: Print individual camera names to ensure no grouping
            print(f"PowerPoint Table - Individual Camera Names: {individual_camera_names}")

            table_info['camera_names'] = individual_camera_names
            table_info['event_types'] = camera_event_pivot.columns.tolist()

            # STEP 10: Create heatmap matrix preserving individual camera data
            heatmap_matrix = []
            max_count = 0

            for camera in individual_camera_names:
                row = []
                for event in table_info['event_types']:
                    count = int(camera_event_pivot.loc[camera, event])
                    row.append(count)
                    max_count = max(max_count, count)
                heatmap_matrix.append(row)

                # DEBUG: Print each camera's data
                total_camera_events = sum(row)
                print(f"PowerPoint Table - {camera}: {total_camera_events} total events")

            table_info['heatmap_matrix'] = heatmap_matrix
            table_info['max_count'] = max_count

            print(f"PowerPoint Table - Final camera count: {len(table_info['camera_names'])} cameras")
        else:
            print("PowerPoint Table - No pivot data available")
            return table_info

    except Exception as e:
        print(f"PowerPoint Table Error: {e}")
        import traceback
        traceback.print_exc()
        return {
            'camera_names': [],
            'event_types': [],
            'heatmap_matrix': [],
            'max_count': 0
        }

    return table_info


def create_event_types_heatmap_table(df, event_data, camera_data):
    """
    Create event types heatmap table data for PowerPoint slide
    Uses EXACTLY the same logic as the dashboard to ensure individual cameras are shown
    """
    # Initialize empty return structure matching what PowerPoint expects
    table_info = {
        'event_types': [],
        'camera_names': [],
        'heatmap_matrix': [],
        'max_count': 0
    }

    if df.empty or not camera_data['camera_column'] or 'Events' not in df.columns:
        return table_info

    camera_column = camera_data['camera_column']

    try:
        # EXACT DASHBOARD LOGIC: Get top 5 event types by count
        top_events = df['Events'].value_counts().head(5)
        top_event_names = top_events.index.tolist()

        # EXACT DASHBOARD LOGIC: Get top 5 cameras by total event count from all data
        top_cameras_all = df[camera_column].value_counts().head(5)
        top_camera_names_all = top_cameras_all.index.tolist()

        if not top_event_names or not top_camera_names_all:
            return table_info

        # EXACT DASHBOARD LOGIC: Create pivot table from ALL data first (not filtering by event types yet)
        full_pivot = df.groupby(['Events', camera_column]).size().unstack(fill_value=0)

        if not full_pivot.empty:
            # EXACT DASHBOARD LOGIC: Filter to only show top 5 event types as rows
            event_camera_pivot = full_pivot.loc[full_pivot.index.intersection(top_event_names)]

            # EXACT DASHBOARD LOGIC: Ensure we have all top 5 events as rows (add missing ones with zeros)
            for event in top_event_names:
                if event not in event_camera_pivot.index:
                    # Add row of zeros for this event
                    event_camera_pivot.loc[event] = 0

            # EXACT DASHBOARD LOGIC: Reorder rows by event frequency (as per top_event_names order)
            event_camera_pivot = event_camera_pivot.reindex(top_event_names)

            # EXACT DASHBOARD LOGIC: Ensure we have all top 5 cameras as columns (add missing ones with zeros)
            for camera in top_camera_names_all:
                if camera not in event_camera_pivot.columns:
                    event_camera_pivot[camera] = 0

            # EXACT DASHBOARD LOGIC: Keep only top 5 cameras and reorder by their total counts
            available_top_cameras = [cam for cam in top_camera_names_all if cam in event_camera_pivot.columns]
            event_camera_pivot = event_camera_pivot[available_top_cameras]

            # Convert to PowerPoint format - preserving EXACT camera names
            table_info['event_types'] = event_camera_pivot.index.tolist()
            table_info['camera_names'] = event_camera_pivot.columns.tolist()  # Individual camera names preserved

            # Create heatmap matrix (2D array) - rows = event types, columns = cameras
            heatmap_matrix = []
            max_count = 0

            for event in table_info['event_types']:
                row = []
                for camera in table_info['camera_names']:
                    count = int(event_camera_pivot.loc[event, camera])
                    row.append(count)
                    max_count = max(max_count, count)
                heatmap_matrix.append(row)

            table_info['heatmap_matrix'] = heatmap_matrix
            table_info['max_count'] = max_count
        else:
            # No data available
            return table_info

    except Exception as e:
        # Return empty structure if any error occurs
        print(f"Error in create_event_types_heatmap_table: {e}")
        table_info = {
            'event_types': [],
            'camera_names': [],
            'heatmap_matrix': [],
            'max_count': 0
        }

    return table_info


def create_events_by_area_chart_for_slide5(df, save_path, camera_data=None):
    """
    Create Events by Area bar chart for slide 5 (right side) - green bars like dashboard
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('white')
    
    # Determine camera column (Area > Camera > None)
    camera_column = None
    if camera_data and camera_data.get('camera_column'):
        camera_column = camera_data['camera_column']
    elif 'Area' in df.columns:
        camera_column = 'Area'
    elif 'Camera' in df.columns:
        camera_column = 'Camera'
    
    if camera_column and not df.empty:
        # Get area/camera counts
        area_counts = df[camera_column].value_counts()
        
        if len(area_counts) > 0:
            # Get top 10 areas only
            top_area_counts = area_counts.head(10)
            
            # Create professional blue bars
            bars = ax.bar(range(len(top_area_counts)), top_area_counts.values, 
                         color=PROFESSIONAL_BLUE, width=0.7, alpha=0.9)
            
            # Add value labels on top of each bar with percentages
            total_events = top_area_counts.sum()
            for i, (bar, count) in enumerate(zip(bars, top_area_counts.values)):
                percentage = (count / total_events * 100) if total_events > 0 else 0
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(top_area_counts.values) * 0.02,
                       f'{count}\n({percentage:.1f}%)', ha='center', va='bottom', 
                       fontweight='bold', fontsize=10, color='#333333')
            
            # Customize the chart
            ax.set_xlabel('Area (Top 10)', fontsize=12, color='#333333', fontweight='bold')
            ax.set_ylabel('Event Count', fontsize=12, color='#333333', fontweight='bold')
            ax.set_title('Events by Area (Top 10)', fontsize=14, color='#333333', fontweight='bold', pad=20)
            
            # Set x-axis labels with proper formatting
            ax.set_xticks(range(len(top_area_counts)))
            labels = []
            for label in top_area_counts.index:
                # Format labels to be more readable
                if len(label) > 12:
                    # Split long labels
                    words = str(label).split()
                    if len(words) > 1:
                        mid = len(words) // 2
                        labels.append('\n'.join([' '.join(words[:mid]), ' '.join(words[mid:])]))
                    else:
                        labels.append(str(label)[:12] + '...')
                else:
                    labels.append(str(label))
            ax.set_xticklabels(labels, fontsize=9, color='#333333', rotation=45, ha='right')
            
            # Style improvements
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#E5E7EB')
            ax.spines['bottom'].set_color('#E5E7EB')
            
            # Set y-axis to start from 0 and add some padding
            ax.set_ylim(0, max(top_area_counts.values) * 1.15)
            
            # Style tick parameters
            ax.tick_params(axis='both', which='major', labelsize=9, colors='#333333')
            
            # Add grid for better readability
            ax.grid(True, alpha=0.3, axis='y')
            ax.set_axisbelow(True)
        else:
            # No areas found
            ax.text(0.5, 0.5, 'No Area data available', ha='center', va='center', 
                   transform=ax.transAxes, fontsize=14)
    else:
        # No camera/area column
        ax.text(0.5, 0.5, 'No Area/Camera data available', ha='center', va='center', 
               transform=ax.transAxes, fontsize=14)
    
    # Save with high DPI for PowerPoint
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()
    
    return save_path


def create_events_by_type_chart_for_slide5(df, save_path, event_data=None):
    """
    Create Events by Type bar chart for slide 5 (left side) - blue bars like dashboard
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor('white')

    # Use accurate event data if provided, otherwise fall back to direct calculation
    if event_data and 'event_counts' in event_data:
        event_counts = event_data['event_counts'].head(10)  # Filter to top 10
    elif 'Events' in df.columns:
        # Get event counts - filter to top 10 only
        event_counts = df['Events'].value_counts().head(10)
    else:
        # No events found
        ax.text(0.5, 0.5, 'No Events data available', ha='center', va='center',
                transform=ax.transAxes, fontsize=14)
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close()
        return save_path

    if len(event_counts) > 0:
        # Create professional blue bars
        bars = ax.bar(range(len(event_counts)), event_counts.values,
                      color=PROFESSIONAL_BLUE, width=0.7, alpha=0.9)

        # Add value labels on top of each bar with percentages
        total_events = event_counts.sum()
        for i, (bar, count) in enumerate(zip(bars, event_counts.values)):
            percentage = (count / total_events * 100) if total_events > 0 else 0
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(event_counts.values) * 0.02,
                    f'{count}\n({percentage:.1f}%)', ha='center', va='bottom',
                    fontweight='bold', fontsize=10, color='#333333')

        # Customize the chart
        ax.set_xlabel('Event Type', fontsize=12, color='#333333', fontweight='bold')
        ax.set_ylabel('Number of Events', fontsize=12, color='#333333', fontweight='bold')
        ax.set_title('Events by Type', fontsize=14, color='#333333', fontweight='bold', pad=20)

        # Set x-axis labels with proper formatting
        ax.set_xticks(range(len(event_counts)))
        labels = []
        for label in event_counts.index:
            # Format labels to be more readable
            if len(label) > 12:
                # Split long labels
                words = label.split()
                if len(words) > 1:
                    mid = len(words) // 2
                    labels.append('\n'.join([' '.join(words[:mid]), ' '.join(words[mid:])]))
                else:
                    labels.append(label[:12] + '...')
            else:
                labels.append(label)
        ax.set_xticklabels(labels, fontsize=9, color='#333333', rotation=45, ha='right')

        # Style improvements
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#E5E7EB')
        ax.spines['bottom'].set_color('#E5E7EB')

        # Set y-axis to start from 0 and add some padding
        ax.set_ylim(0, max(event_counts.values) * 1.15)

        # Style tick parameters
        ax.tick_params(axis='both', which='major', labelsize=9, colors='#333333')

        # Add grid for better readability
        ax.grid(True, alpha=0.3, axis='y')
        ax.set_axisbelow(True)
    else:
        # No events found
        ax.text(0.5, 0.5, 'No Events data available', ha='center', va='center',
                transform=ax.transAxes, fontsize=14)

    # Save with high DPI for PowerPoint
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def create_events_by_severity_pie_for_slide6(df, save_path, severity_data=None):
    """
    Create clean Events by Severity pie chart for slide 6 (no external labels)
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 8))
    fig.patch.set_facecolor('white')

    # Use accurate severity data if provided, otherwise fall back to direct calculation
    if severity_data and 'severity_counts' in severity_data:
        severity_counts = severity_data['severity_counts']
    else:
        # Get severity counts using the same logic as dashboard
        severity_counts = pd.Series(dtype=int)
        if 'Severity' in df.columns and not df.empty:
            severity_counts = df['Severity'].value_counts()

    if len(severity_counts) > 0:
        # Use minimal professional colors for pie charts
        colors = PIE_COLORS[:len(severity_counts)]
        if len(severity_counts) > len(PIE_COLORS):
            colors.extend(['#94A3B8'] * (len(severity_counts) - len(PIE_COLORS)))

        # Custom autopct function to show percentages only for top 5
        def autopct_func(pct, allvalues):
            absolute = int(pct / 100. * sum(allvalues))
            # Get the index of this slice
            cumsum = 0
            for i, val in enumerate(allvalues):
                cumsum += val
                if cumsum >= absolute:
                    return f'{pct:.1f}%' if i < 5 else ''  # Show percentage only for top 5
            return ''

        # Create pie chart WITHOUT external labels
        wedges, texts, autotexts = ax.pie(severity_counts.values,
                                          labels=None,  # No external labels
                                          colors=colors[:len(severity_counts)],
                                          autopct=lambda pct: autopct_func(pct, severity_counts.values),
                                          startangle=90,
                                          textprops={'fontsize': 12, 'fontweight': 'bold'})

        # Customize autopct text
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(12)

        # Add title
        ax.set_title('Events by Severity', fontsize=16, fontweight='bold', pad=20, color='#333333')
        ax.axis('equal')

    else:
        # No severity data found
        ax.text(0.5, 0.5, 'No Severity data available', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')
        ax.set_title('Events by Severity', fontsize=16, fontweight='bold', pad=20, color='#333333')

    # Save with high DPI for PowerPoint
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def create_events_by_type_pie_for_slide6(df, save_path, event_data=None):
    """
    Create clean Events by Type pie chart for slide 6 (no external labels)
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 8))
    fig.patch.set_facecolor('white')

    # Use accurate event data if provided, otherwise fall back to direct calculation
    if event_data and 'event_counts' in event_data:
        event_counts = event_data['event_counts']
    elif 'Events' in df.columns:
        # Get event counts - all event types
        event_counts = df['Events'].value_counts()
    else:
        # No events found
        event_counts = pd.Series(dtype=int)

    if len(event_counts) > 0:
        # Use minimal professional colors for pie charts
        colors = PIE_COLORS[:len(event_counts)]
        if len(event_counts) > len(PIE_COLORS):
            colors.extend(['#94A3B8'] * (len(event_counts) - len(PIE_COLORS)))

        # Custom autopct function to show percentages only for top 5
        def autopct_func(pct, allvalues):
            absolute = int(pct / 100. * sum(allvalues))
            # Get the index of this slice
            cumsum = 0
            for i, val in enumerate(allvalues):
                cumsum += val
                if cumsum >= absolute:
                    return f'{pct:.1f}%' if i < 5 else ''  # Show percentage only for top 5
            return ''

        # Create pie chart WITHOUT external labels
        wedges, texts, autotexts = ax.pie(event_counts.values,
                                          labels=None,  # No external labels
                                          colors=colors[:len(event_counts)],
                                          autopct=lambda pct: autopct_func(pct, event_counts.values),
                                          startangle=90,
                                          textprops={'fontsize': 12, 'fontweight': 'bold'})

        # Customize autopct text
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(12)

        # Add title
        ax.set_title('Events by Type', fontsize=16, fontweight='bold', pad=20, color='#333333')
        ax.axis('equal')

    else:
        # No event data found
        ax.text(0.5, 0.5, 'No Events data available', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')
        ax.set_title('Events by Type', fontsize=16, fontweight='bold', pad=20, color='#333333')

    # Save with high DPI for PowerPoint
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def create_severity_summary_table_for_slide6(slide, df, severity_data):
    """
    Create severity summary table below the severity pie chart on slide 6
    """
    # Get severity data
    if severity_data and 'severity_breakdown' in severity_data and len(severity_data['severity_breakdown']) > 0:
        severity_counts = severity_data['severity_breakdown']
    elif 'Severity' in df.columns and not df.empty:
        severity_counts = df['Severity'].value_counts()
    else:
        return  # No severity data to show

    total_events = severity_counts.sum()

    # Prepare table data
    table_data = []
    table_data.append(['Severity Level', 'Count', 'Percentage'])  # Header

    for severity, count in severity_counts.items():
        percentage = (count / total_events) * 100
        table_data.append([str(severity), str(count), f'{percentage:.1f}%'])

    # Create table
    rows = len(table_data)
    cols = 3

    # Position table under the severity pie chart (left side) with proper margins
    table = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(4.8), Inches(5.5), Inches(2.5))

    # Set table data and styling
    for i, row_data in enumerate(table_data):
        for j, cell_data in enumerate(row_data):
            cell = table.table.cell(i, j)
            cell.text = cell_data

            # Header styling
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(74, 85, 104)  # Dark blue-gray
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
                    paragraph.font.size = Pt(11)
                    paragraph.alignment = PP_ALIGN.LEFT
                    paragraph.font.name = "Arial"
            else:
                # Data row styling with severity colors
                severity_name = severity_counts.index[i - 1].lower()
                if 'low' in severity_name:
                    color = RGBColor(198, 246, 213)  # Light green
                elif 'moderate' in severity_name:
                    color = RGBColor(254, 215, 170)  # Light orange
                elif 'high' in severity_name:
                    color = RGBColor(254, 178, 178)  # Light red
                else:
                    color = RGBColor(226, 232, 240)  # Light gray

                cell.fill.solid()
                cell.fill.fore_color.rgb = color
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor(51, 65, 85)
                    paragraph.font.size = Pt(10)
                    paragraph.alignment = PP_ALIGN.LEFT
                    paragraph.font.name = "Arial"

            # Vertical alignment
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE


def create_events_summary_table_for_slide6(slide, df, event_data):
    """
    Create events summary table below the events pie chart on slide 6
    """
    # Get event data
    if event_data and 'event_counts' in event_data:
        event_counts = event_data['event_counts']
    elif 'Events' in df.columns and not df.empty:
        event_counts = df['Events'].value_counts()
    else:
        return  # No event data to show

    total_events = event_counts.sum()

    # Get top 8 event types for the table (changed from default to top 8)
    top_8_events = event_counts.head(8)

    # Prepare table data
    table_data = []
    table_data.append(['Event Type', 'Count', 'Percentage'])  # Header

    for event_type, count in top_8_events.items():
        percentage = (count / total_events) * 100
        # Truncate long event type names
        display_name = str(event_type)[:25] + '...' if len(str(event_type)) > 25 else str(event_type)
        table_data.append([display_name, str(count), f'{percentage:.1f}%'])

    # Create table
    rows = len(table_data)
    cols = 3

    # Position table under the events pie chart (right side) with proper margins
    table = slide.shapes.add_table(rows, cols, Inches(7.0), Inches(4.8), Inches(5.5), Inches(2.5))

    # Use minimal professional colors
    colors = PIE_COLORS

    # Set table data and styling
    for i, row_data in enumerate(table_data):
        for j, cell_data in enumerate(row_data):
            cell = table.table.cell(i, j)
            cell.text = cell_data

            # Header styling
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(99, 102, 241)  # Purple header
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
                    paragraph.font.size = Pt(11)
                    paragraph.alignment = PP_ALIGN.LEFT
                    paragraph.font.name = "Arial"
            else:
                # Data row styling with event type colors
                color_idx = (i - 1) % len(colors)
                light_color = plt.matplotlib.colors.to_rgba(colors[color_idx], alpha=0.3)
                rgb_color = tuple(int(c * 255) for c in light_color[:3])

                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(*rgb_color)
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor(51, 65, 85)
                    paragraph.font.size = Pt(10)
                    paragraph.alignment = PP_ALIGN.LEFT
                    paragraph.font.name = "Arial"

            # Vertical alignment
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE


def create_areas_pie_chart_for_slide7(df, save_path):
    """
    Create "Number of Events per Area" pie chart with percentage labels for top 5 areas only
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 8))
    fig.patch.set_facecolor('white')

    # Get area data using the centralized function
    camera_data = get_accurate_camera_data(df)

    if camera_data and 'camera_counts' in camera_data:
        area_counts = camera_data['camera_counts']

        if len(area_counts) > 0:
            # Define colors similar to your image
            colors = PIE_COLORS[:len(area_counts)]
        if len(area_counts) > len(PIE_COLORS):
            colors.extend(['#94A3B8'] * (len(area_counts) - len(PIE_COLORS)))

            # Custom autopct function - show percentage only for top 5 areas
            def autopct_func(pct, allvalues):
                # Calculate which slice this is by percentage
                sorted_values = sorted(allvalues, reverse=True)
                current_value = pct * sum(allvalues) / 100

                # Check if this value is in top 5
                top_5_threshold = sorted_values[4] if len(sorted_values) >= 5 else 0

                if current_value >= top_5_threshold and len([v for v in sorted_values if v >= current_value]) <= 5:
                    return f'{pct:.1f}%'
                else:
                    return ''

            # Create pie chart
            wedges, texts, autotexts = ax.pie(area_counts.values,
                                              autopct=lambda pct: autopct_func(pct, area_counts.values),
                                              startangle=90,
                                              colors=colors[:len(area_counts)],
                                              textprops={'fontsize': 12, 'fontweight': 'bold'})

            # Customize autopct text
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(12)

            ax.set_title('Number of Events per Area', fontsize=16, fontweight='bold',
                         pad=20, color='#333333')
            ax.axis('equal')
        else:
            ax.text(0.5, 0.5, 'No Area data available', ha='center', va='center',
                    transform=ax.transAxes, fontsize=14, color='#333333')
            ax.set_title('Number of Events per Area', fontsize=16, fontweight='bold',
                         pad=20, color='#333333')
    else:
        ax.text(0.5, 0.5, 'No Area data available', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')
        ax.set_title('Number of Events per Area', fontsize=16, fontweight='bold',
                     pad=20, color='#333333')

    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def create_areas_summary_table_for_slide7(slide, df):
    """
    Create summary table for top 10 areas under the pie chart
    """
    # Get area data
    camera_data = get_accurate_camera_data(df)

    if camera_data and 'camera_counts' in camera_data:
        area_counts = camera_data['camera_counts']
        total_events = area_counts.sum()

        # Get top 8 areas
        top_8_areas = area_counts.head(8)

        # Use minimal professional colors
        colors = PIE_COLORS

        # Prepare table data
        table_data = []
        table_data.append(['Area', 'Count', 'Percentage'])  # Header

        for i, (area, count) in enumerate(top_8_areas.items()):
            percentage = (count / total_events) * 100
            table_data.append([str(area), str(count), f'{percentage:.1f}%'])

        # Create table
        rows = len(table_data)
        cols = 3

        # Position table under the pie chart with proper margins
        table = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(4.5), Inches(5.5), Inches(2.8))

        # Set table data and styling
        for i, row_data in enumerate(table_data):
            for j, cell_data in enumerate(row_data):
                cell = table.table.cell(i, j)
                cell.text = cell_data

                # Header styling
                if i == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(74, 85, 104)  # Dark blue-gray
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.bold = True
                        paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
                        paragraph.font.size = Pt(11)
                        paragraph.alignment = PP_ALIGN.LEFT
                        paragraph.font.name = "Arial"
                else:
                    # Data row styling with matching colors
                    color_idx = (i - 1) % len(colors)
                    light_color = plt.matplotlib.colors.to_rgba(colors[color_idx], alpha=0.3)
                    rgb_color = tuple(int(c * 255) for c in light_color[:3])

                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(*rgb_color)
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.bold = True
                        paragraph.font.color.rgb = RGBColor(51, 65, 85)
                        paragraph.font.size = Pt(10)
                        paragraph.alignment = PP_ALIGN.LEFT
                        paragraph.font.name = "Arial"

                # Vertical alignment
                cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE


def create_top_three_areas_cards_for_slide7(slide, df):
    """
    Create top three areas cards on the right side of slide 7
    """
    # Get area data
    camera_data = get_accurate_camera_data(df)

    if camera_data and 'camera_counts' in camera_data:
        area_counts = camera_data['camera_counts']
        total_events = area_counts.sum()

        # Get top 3 areas
        top_3_areas = area_counts.head(3)

        # Define card colors and positions
        card_colors = [
            RGBColor(255, 235, 59),   # Yellow for #1
            RGBColor(169, 169, 169),  # Gray for #2
            RGBColor(255, 152, 0)     # Orange for #3
        ]

        card_y_positions = [0.9, 2.7, 4.5]  # Vertical positions for 3 cards

        for i, (area, count) in enumerate(top_3_areas.items()):
            if i >= 3:  # Only show top 3
                break

            percentage = (count / total_events) * 100

            # Calculate near miss events for this area
            if camera_data['camera_column'] in df.columns:
                area_df = df[df[camera_data['camera_column']] == area]
                near_miss_count = 0
                if 'Events' in area_df.columns:
                    near_miss_pattern = r'near[-\s]*miss'
                    near_miss_count = area_df['Events'].str.contains(
                        near_miss_pattern, case=False, na=False, regex=True).sum()
            else:
                near_miss_count = 0

            # Create card background
            card = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(7), Inches(card_y_positions[i] + 0.1), Inches(5.5), Inches(1.2))
            card.fill.solid()
            card.fill.fore_color.rgb = card_colors[i]
            card.line.color.rgb = RGBColor(51, 65, 85)
            card.line.width = Pt(1)
            
            # Create rank circle
            rank_circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(7.1), Inches(card_y_positions[i] + 0.2), Inches(0.3), Inches(0.3))
            rank_circle.fill.solid()
            rank_circle.fill.fore_color.rgb = RGBColor(51, 65, 85)  # Dark blue-gray
            rank_circle.line.fill.background()

            # Add rank number
            rank_text = slide.shapes.add_textbox(Inches(7.3), Inches(
                card_y_positions[i] + 0.15), Inches(0.2), Inches(0.3))
            rank_text_frame = rank_text.text_frame
            rank_text_frame.text = f"#{i+1}"
            rank_text_frame.paragraphs[0].font.size = Pt(11)
            rank_text_frame.paragraphs[0].font.bold = True
            rank_text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)  # White text
            rank_text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            rank_text_frame.paragraphs[0].font.name = "Arial"

            # Add area name
            area_name_box = slide.shapes.add_textbox(Inches(7.7), Inches(
                card_y_positions[i] + 0.1), Inches(4.5), Inches(0.4))
            area_name_frame = area_name_box.text_frame
            area_name_frame.text = str(area)
            area_name_frame.paragraphs[0].font.size = Pt(16)
            area_name_frame.paragraphs[0].font.bold = True
            area_name_frame.paragraphs[0].font.color.rgb = RGBColor(51, 65, 85)
            area_name_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            area_name_frame.paragraphs[0].font.name = "Arial"

            # Add statistics
            stats_box = slide.shapes.add_textbox(Inches(7.2), Inches(card_y_positions[i] + 0.6), Inches(5), Inches(0.8))
            stats_frame = stats_box.text_frame
            stats_frame.text = f"📊 Total Events: {count}\n📈 Percentage: {percentage:.1f}%\n⚠️ Near Miss: {near_miss_count}"

            for paragraph in stats_frame.paragraphs:
                paragraph.font.size = Pt(11)
                paragraph.font.bold = True
                paragraph.font.color.rgb = RGBColor(51, 65, 85)
                paragraph.alignment = PP_ALIGN.LEFT
                paragraph.font.name = "Arial"


def create_camera_events_pie_chart_for_slide8(area_df, area_name, save_path):
    """
    Create Camera Distribution pie chart for slide 8 - exactly like dashboard
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 8))
    fig.patch.set_facecolor('white')

    if area_df.empty:
        ax.text(0.5, 0.5, f'No data available for {area_name}', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')
        ax.set_title(f'Camera Distribution - {area_name}', fontsize=16, fontweight='bold',
                     pad=20, color='#333333')
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close()
        return save_path

    # Use same logic as dashboard - Camera column or Area column
    camera_column = 'Camera' if 'Camera' in area_df.columns else 'Area'

    if camera_column in area_df.columns:
        camera_counts = area_df[camera_column].value_counts()

        if len(camera_counts) > 0:
            total_events = camera_counts.sum()

            # Colors similar to dashboard
            colors = PIE_COLORS  # Red, Teal, Blue

            # Custom autopct function to show percentages only for top 5
            def autopct_func(pct, allvalues):
                absolute = int(pct / 100. * sum(allvalues))
                # Get the index of this slice
                cumsum = 0
                for i, val in enumerate(allvalues):
                    cumsum += val
                    if cumsum >= absolute:
                        return f'{pct:.1f}%' if i < 5 else ''  # Show percentage only for top 5
                return ''

            # Create pie chart
            wedges, texts, autotexts = ax.pie(camera_counts.values,
                                              labels=None,
                                              autopct=lambda pct: autopct_func(pct, camera_counts.values),
                                              startangle=90,
                                              colors=colors[:len(camera_counts)],
                                              textprops={'fontsize': 12, 'fontweight': 'bold'})

            # Customize autopct text
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(12)

            # Create legend
            legend_labels = [f"{camera}: {count}" for camera, count in camera_counts.items()]
            ax.legend(wedges, legend_labels, title="Distribution", loc="center left",
                      bbox_to_anchor=(1, 0, 0.5, 1), fontsize=10)

            ax.set_title(f'Camera Distribution - {area_name}', fontsize=16, fontweight='bold',
                         pad=20, color='#333333')
            ax.axis('equal')
        else:
            ax.text(0.5, 0.5, f'No camera data available for {area_name}', ha='center', va='center',
                    transform=ax.transAxes, fontsize=14, color='#333333')
            ax.set_title(f'Camera Distribution - {area_name}', fontsize=16, fontweight='bold',
                         pad=20, color='#333333')
    else:
        ax.text(0.5, 0.5, f'No camera data available for {area_name}', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')
        ax.set_title(f'Camera Distribution - {area_name}', fontsize=16, fontweight='bold',
                     pad=20, color='#333333')

    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def create_camera_severity_pie_chart_for_slide8(area_df, area_name, save_path):
    """
    Create Severity Distribution pie chart for slide 8 - exactly like dashboard
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(10, 8))
    fig.patch.set_facecolor('white')

    if area_df.empty or 'Severity' not in area_df.columns:
        ax.text(0.5, 0.5, f'No severity data available for {area_name}', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')
        ax.set_title(f'Severity Distribution - {area_name}', fontsize=16, fontweight='bold',
                     pad=20, color='#333333')
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close()
        return save_path

    severity_counts = area_df['Severity'].value_counts()

    if len(severity_counts) > 0:
        total_events = severity_counts.sum()

        # Colors similar to dashboard - severity based
        severity_colors = {
            'Moderate': '#FF6B6B',   # Red
            'High': '#FF8C42',       # Orange
            'Low': '#FFD93D',        # Yellow
            'Critical': '#DC143C'    # Dark Red
        }

        # Map colors to actual severity levels in data
        colors = PIE_COLORS

        # Custom autopct function to show percentages only for top 5
        def autopct_func(pct, allvalues):
            absolute = int(pct / 100. * sum(allvalues))
            # Get the index of this slice
            cumsum = 0
            for i, val in enumerate(allvalues):
                cumsum += val
                if cumsum >= absolute:
                    return f'{pct:.1f}%' if i < 5 else ''  # Show percentage only for top 5
            return ''

        # Create pie chart
        wedges, texts, autotexts = ax.pie(severity_counts.values,
                                          labels=None,
                                          autopct=lambda pct: autopct_func(pct, severity_counts.values),
                                          startangle=90,
                                          colors=colors,
                                          textprops={'fontsize': 12, 'fontweight': 'bold'})

        # Customize autopct text
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(12)

        # Create legend
        legend_labels = [f"{severity}: {count}" for severity, count in severity_counts.items()]
        ax.legend(wedges, legend_labels, title="Severity Levels", loc="center left",
                  bbox_to_anchor=(1, 0, 0.5, 1), fontsize=10)

        ax.set_title(f'Severity Distribution - {area_name}', fontsize=16, fontweight='bold',
                     pad=20, color='#333333')
        ax.axis('equal')
    else:
        ax.text(0.5, 0.5, f'No severity data available for {area_name}', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')
        ax.set_title(f'Severity Distribution - {area_name}', fontsize=16, fontweight='bold',
                     pad=20, color='#333333')

    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def plot_excel_comparison_chart(df1, df2, label1="Dataset 1", label2="Dataset 2"):
    """
    Create a comparison chart between two Excel datasets showing average fluctuation in events
    with different colored lines and percentage reduction calculation.
    """
    try:
        # Process both datasets to get daily event counts
        def process_dataset(df, dataset_name):
            if df is None or df.empty:
                return None, None, None

            # Make a copy to avoid modifying original
            df_copy = df.copy()

            # Try to find a date column
            date_col = None
            for col in df_copy.columns:
                if any(keyword in col.lower() for keyword in ['date', 'time', 'created', 'timestamp']):
                    date_col = col
                    break

            if date_col is None:
                st.error(
                    f"No date column found in {dataset_name}. Please ensure your Excel file has a date/time column.")
                return None, None, None

            # Convert to datetime
            df_copy[date_col] = pd.to_datetime(df_copy[date_col], dayfirst=True, errors='coerce')
            df_copy = df_copy.dropna(subset=[date_col])

            if df_copy.empty:
                st.error(f"No valid dates found in {dataset_name}")
                return None, None, None

            # Extract date only (remove time component) - handle both datetime and date objects
            if hasattr(df_copy[date_col].iloc[0], 'date'):
                df_copy['Date'] = df_copy[date_col].dt.date
            else:
                df_copy['Date'] = df_copy[date_col]

            # Count events per day
            daily_counts = df_copy.groupby('Date').size().reset_index(name='Events')
            daily_counts['Date'] = pd.to_datetime(daily_counts['Date'])

            return daily_counts, df_copy[date_col].min(), df_copy[date_col].max()

        # Process both datasets
        daily1, min_date1, max_date1 = process_dataset(df1, label1)
        daily2, min_date2, max_date2 = process_dataset(df2, label2)

        if daily1 is None or daily2 is None:
            return None

        # Handle date range creation more carefully
        if hasattr(min_date1, 'date'):
            start_date = min(min_date1.date(), min_date2.date())
            end_date = max(max_date1.date(), max_date2.date())
        else:
            start_date = min(min_date1, min_date2)
            end_date = max(max_date1, max_date2)
            if hasattr(start_date, 'date'):
                start_date = start_date.date()
                end_date = end_date.date()

        # Extract day-of-month and month names for overlay comparison
        daily1['Day'] = daily1['Date'].dt.day
        daily1['Month'] = daily1['Date'].dt.month
        daily1['Year'] = daily1['Date'].dt.year
        daily1['MonthName'] = daily1['Date'].dt.strftime('%B')
        
        daily2['Day'] = daily2['Date'].dt.day
        daily2['Month'] = daily2['Date'].dt.month
        daily2['Year'] = daily2['Date'].dt.year
        daily2['MonthName'] = daily2['Date'].dt.strftime('%B')

        # Get month names for labels
        month1_name = daily1['MonthName'].iloc[0] if not daily1.empty else "Month 1"
        month2_name = daily2['MonthName'].iloc[0] if not daily2.empty else "Month 2"
        year1 = daily1['Year'].iloc[0] if not daily1.empty else 2024
        year2 = daily2['Year'].iloc[0] if not daily2.empty else 2024

        # Create day-wise comparison (1-31 days)
        days_range = range(1, 32)  # Days 1-31
        
        # Create arrays for plotting - one for each month
        month1_events = []
        month2_events = []
        
        for day in days_range:
            # Get events for this day in month 1
            day_events1 = daily1[daily1['Day'] == day]['Events'].sum()
            month1_events.append(day_events1)
            
            # Get events for this day in month 2
            day_events2 = daily2[daily2['Day'] == day]['Events'].sum()
            month2_events.append(day_events2)

        # Calculate averages
        avg1 = sum(month1_events) / len([x for x in month1_events if x > 0]) if any(x > 0 for x in month1_events) else 0
        avg2 = sum(month2_events) / len([x for x in month2_events if x > 0]) if any(x > 0 for x in month2_events) else 0

        # Calculate percentage change
        if avg1 > 0:
            percentage_change = ((avg2 - avg1) / avg1) * 100
        else:
            percentage_change = 0

        # Create the plot
        fig, ax = plt.subplots(figsize=(12, 5))

        # Plot both months with different colors
        ax.plot(days_range, month1_events,
                linewidth=2.5, color='#DC2626', alpha=0.8, label=f'{month1_name} {year1}', marker='o', markersize=5)
        ax.plot(days_range, month2_events,
                linewidth=2.5, color='#7C3AED', alpha=0.8, label=f'{month2_name} {year2}', marker='o', markersize=5)

        # Event count labels removed as requested

        # Add average lines (dashed)
        if avg1 > 0:
            ax.axhline(y=avg1, color='#DC2626', linestyle='--', alpha=0.6, linewidth=1.5)
        if avg2 > 0:
            ax.axhline(y=avg2, color='#7C3AED', linestyle='--', alpha=0.6, linewidth=1.5)

        # Add percentage change text box
        bbox_props = dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.8, edgecolor='gray')
        change_text = f"Average Fluctuation in Events\n{percentage_change:+.2f}%"
        ax.text(0.02, 0.98, change_text, transform=ax.transAxes, fontsize=12, 
                verticalalignment='top', bbox=bbox_props, fontweight='bold')

        # Set x-axis to show days 1-31
        ax.set_xticks(range(1, 32))  # Show all days 1-31
        ax.set_xticklabels(range(1, 32), rotation=45, ha='right', fontsize=9)

        # Add legend
        ax.legend()

        # Style the plot to match your example
        ax.set_title('📈 Average Fluctuation in Events',
                     fontsize=14, fontweight='bold')
        ax.set_xlabel('Date', fontsize=11)
        ax.set_ylabel('Number of Events', fontsize=11)
        
        # Add grid like in your example
        ax.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)
        ax.set_facecolor('#FAFAFA')
        
        # Set y-axis to start slightly above 0 for better comparison
        ax.set_ylim(bottom=-5)

        # Add tight layout
        plt.tight_layout()

        # Convert to base64
        buffer = BytesIO()
        plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight', facecolor='white')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.getvalue()).decode()
        plt.close()

        return image_base64, avg1, avg2, percentage_change

    except Exception as e:
        st.error(f"Error creating comparison chart: {str(e)}")
        return None


def create_events_per_day_chart_for_slide1(df, save_path, start_date, end_date):
    """
    Create Events per Day trend chart for slide 1 using the same logic as the dashboard
    """
    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(12, 6))
    fig.patch.set_facecolor('white')

    # Validate data before processing
    if df.empty or 'Date' not in df.columns:
        ax.text(0.5, 0.5, 'No date data available', ha='center', va='center',
                transform=ax.transAxes, fontsize=14)
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close()
        return save_path

    # Group by date and count events
    events_per_day = df.groupby(df['Date'].dt.date)['Events'].count().reset_index()
    events_per_day.columns = ['Date', 'Event_Count']
    
    # Create complete date range
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    complete_dates = pd.DataFrame({'Date': date_range.date})
    complete_events = complete_dates.merge(events_per_day, on='Date', how='left')
    complete_events['Event_Count'] = complete_events['Event_Count'].fillna(0)

    if len(complete_events) > 0:
        complete_events = complete_events.sort_values('Date')
        x = range(len(complete_events))

        # Plot line connecting all points (including zeros)
        ax.plot(x, complete_events['Event_Count'], linewidth=3, color='#8B5CF6', alpha=0.8)
        
        # Add scatter points for non-zero events only
        non_zero_mask = complete_events['Event_Count'] > 0
        non_zero_positions = [i for i, is_nonzero in enumerate(non_zero_mask) if is_nonzero]
        non_zero_counts = complete_events[complete_events['Event_Count'] > 0]['Event_Count']
        
        if len(non_zero_positions) > 0:
            ax.scatter(non_zero_positions, non_zero_counts, 
                      color='#8B5CF6', s=80, zorder=5, label='With Events')
            
            # Add value annotations for non-zero points only
            for pos, count in zip(non_zero_positions, non_zero_counts):
                ax.annotate(str(int(count)), (pos, count), textcoords="offset points",
                           xytext=(0, 10), ha='center', fontsize=10, 
                           fontweight='bold', color='#333333')

        # Use smart date display logic
        date_labels = [d.strftime('%Y-%m-%d') for d in complete_events['Date']]
        positions_to_show, labels_to_show = get_smart_date_display(date_labels, max_labels=30)

        ax.set_xticks(positions_to_show)
        ax.set_xticklabels(labels_to_show, rotation=45, ha='right', fontsize=9, color='#333333')
        ax.legend(fontsize=10)

        # Style improvements
        ax.set_xlabel('Date', fontsize=12, color='#333333', fontweight='bold')
        ax.set_ylabel('Number of Events', fontsize=12, color='#333333', fontweight='bold')

        # Set y-axis to start slightly above 0 and add some padding
        max_count = max(complete_events['Event_Count']) if len(complete_events) > 0 else 1
        ax.set_ylim(-2, max_count * 1.15)

        # Style the chart
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#E5E7EB')
        ax.spines['bottom'].set_color('#E5E7EB')

        # Style tick parameters
        ax.tick_params(axis='both', which='major', labelsize=10, colors='#333333')

        # Add grid for better readability
        ax.grid(True, alpha=0.3, axis='y')
        ax.set_axisbelow(True)
    else:
        # No events with data
        ax.text(0.5, 0.5, 'No events found for the selected period', ha='center', va='center',
                transform=ax.transAxes, fontsize=14, color='#333333')

    # Save with high DPI for PowerPoint
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close()

    return save_path


def create_comparison_chart(df1, df2, label1, label2, save_path):
    """
    Create a comparison chart between two datasets
    """
    try:
        # Convert Date column to datetime if it's not already
        if 'Date' in df1.columns:
            df1['Date'] = pd.to_datetime(df1['Date'])
        if 'Date' in df2.columns:
            df2['Date'] = pd.to_datetime(df2['Date'])

        # Group by date and count events
        daily1 = df1.groupby('Date').size().reset_index(name='Events')
        daily2 = df2.groupby('Date').size().reset_index(name='Events')

        # Get date range
        min_date1 = daily1['Date'].min() if not daily1.empty else pd.Timestamp.now()
        max_date1 = daily1['Date'].max() if not daily1.empty else pd.Timestamp.now()
        min_date2 = daily2['Date'].min() if not daily2.empty else pd.Timestamp.now()
        max_date2 = daily2['Date'].max() if not daily2.empty else pd.Timestamp.now()

        # Create complete date range
        start_date = min(min_date1, min_date2)
        end_date = max(max_date1, max_date2)
        if hasattr(start_date, 'date'):
            start_date = start_date.date()
            end_date = end_date.date()

        # Create complete date range
        date_range = pd.date_range(start=start_date, end=end_date, freq='D')
        complete_df = pd.DataFrame({'Date': date_range})

        # Merge with daily counts (fill missing dates with 0)
        daily1_complete = complete_df.merge(daily1, on='Date', how='left').fillna(0)
        daily2_complete = complete_df.merge(daily2, on='Date', how='left').fillna(0)

        # Calculate averages
        avg1 = daily1_complete['Events'].mean()
        avg2 = daily2_complete['Events'].mean()

        # Calculate percentage change
        if avg1 > 0:
            percentage_change = ((avg2 - avg1) / avg1) * 100
        else:
            percentage_change = 0

        # Create the plot with professional styling for PowerPoint
        fig, ax = plt.subplots(figsize=(12, 5))

        # Show all dates like in the dashboard
        x_positions = range(len(daily1_complete))
        date_labels = [d.strftime('%Y-%m-%d') for d in daily1_complete['Date']]

        # Plot lines connecting all points (including zeros)
        ax.plot(x_positions, daily1_complete['Events'],
                linewidth=3, color='#FF6B6B', alpha=0.8, label=f'{label1}')
        ax.plot(x_positions, daily2_complete['Events'],
                linewidth=3, color='#4ECDC4', alpha=0.8, label=f'{label2}')

        # Add scatter points for all dates (including zeros)
        ax.scatter(x_positions, daily1_complete['Events'],
                   color='#FF6B6B', s=80, zorder=5, alpha=0.9)
        ax.scatter(x_positions, daily2_complete['Events'],
                   color='#4ECDC4', s=80, zorder=5, alpha=0.9)

        # Set x-axis with Smart Date Display Logic
        if len(date_labels) <= 7:  # Week or less - show all
            show_x = list(x_positions)
            show_labels = date_labels
        elif len(date_labels) <= 31:  # Month or less - show all
            show_x = list(x_positions)
            show_labels = date_labels
        else:  # Longer periods - smart skipping
            if len(date_labels) <= 60:  # 2 months - show every 2nd
                step = 2
            elif len(date_labels) <= 180:  # 6 months - show every 7th
                step = 7
            else:  # Longer periods - show every 14th
                step = 14
            show_x = list(x_positions)[::step]
            show_labels = date_labels[::step]

        ax.set_xticks(show_x)
        ax.set_xticklabels(show_labels, rotation=45, ha='right', fontsize=10)

        # Style the plot for PowerPoint
        ax.set_title('📈 Average Fluctuation in Events',
                     fontsize=16, fontweight='bold', color='#1f2937', pad=20)
        ax.set_xlabel('Date', fontsize=12, fontweight='bold')
        ax.set_ylabel('Number of Events', fontsize=12, fontweight='bold')

        # Add legend with better styling
        ax.legend(loc='upper right', frameon=True, fancybox=True, shadow=True, fontsize=11)

        # Add grid for better readability
        ax.grid(True, alpha=0.3, linestyle='--')

        # Set background color
        ax.set_facecolor('#fafafa')
        fig.patch.set_facecolor('white')

        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()

        return save_path, avg1, avg2, percentage_change

    except Exception as e:
        # Create error chart
        fig, ax = plt.subplots(figsize=(12, 5))
        ax.text(0.5, 0.5, f'Error creating comparison chart: {str(e)}', ha='center', va='center',
                transform=ax.transAxes, fontsize=12, color='red')
        ax.set_title('📈 Average Fluctuation in Events',
                     fontsize=14, fontweight='bold')
        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        return save_path, 0, 0, 0


if __name__ == "__main__":
    main()
