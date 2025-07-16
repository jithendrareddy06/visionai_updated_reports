import os
import gc
import shutil
import tempfile
import psutil
from io import BytesIO
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import calendar
from PIL import Image

class PowerPointGenerator:
    def __init__(self, min_memory_gb=2):
        """Initialize PowerPoint generator with minimum memory requirement"""
        self.min_memory_gb = min_memory_gb
        self.temp_dir = None
        self.prs = None
        self.cleanup_performed = False

    def __enter__(self):
        """Context manager entry"""
        self._check_system_resources()
        self._create_temp_directory()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - ensures cleanup"""
        self.cleanup()
        
    def cleanup(self):
        """Clean up resources"""
        if not self.cleanup_performed:
            plt.close('all')
            gc.collect()
            
            if self.temp_dir and os.path.exists(self.temp_dir):
                try:
                    shutil.rmtree(self.temp_dir)
                except Exception as e:
                    print(f"Warning: Failed to remove temporary directory: {e}")
            
            self.cleanup_performed = True

    def _check_system_resources(self):
        """Check if system has enough resources"""
        available_memory = psutil.virtual_memory().available / (1024 * 1024 * 1024)  # GB
        if available_memory < self.min_memory_gb:
            raise ValueError(
                f"Insufficient memory available ({available_memory:.1f}GB). "
                f"Need at least {self.min_memory_gb}GB. Close other applications or reduce dataset size."
            )

    def _create_temp_directory(self):
        """Create temporary directory for chart images"""
        try:
            self.temp_dir = tempfile.mkdtemp()
            if not os.path.exists(self.temp_dir):
                raise ValueError("Failed to create temporary directory")
        except Exception as e:
            raise ValueError(f"Error creating temporary directory: {str(e)}. Check disk space and permissions.")

    def validate_data(self, df):
        """Validate input dataframe"""
        if df is None or df.empty:
            raise ValueError("No data available. Please ensure your data is not empty.")

        # Check required columns
        required_columns = ['Date', 'Events']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(
                f"Missing required columns: {', '.join(missing_columns)}. "
                "Please ensure your data has all required columns."
            )

        # Validate Date column
        if not pd.api.types.is_datetime64_any_dtype(df['Date']):
            try:
                df['Date'] = pd.to_datetime(df['Date'])
            except Exception as e:
                raise ValueError("Invalid date format in Date column. Please ensure dates are in a valid format.") from e

        # Check data size
        if len(df) > 10000:
            print(f"Warning: Large dataset detected ({len(df)} rows). This may impact performance.")

        # Check date range
        date_range = (df['Date'].max() - df['Date'].min()).days
        if date_range > 365:
            print(f"Warning: Data spans {date_range} days. Consider using a smaller date range for better analysis.")

        # Check unique event types
        event_types = df['Events'].nunique()
        if event_types > 50:
            print(f"Warning: Large number of unique event types ({event_types}). This may affect chart readability.")

        return df

    def initialize_presentation(self, template=None):
        """Initialize PowerPoint presentation with proper error handling"""
        try:
            self.prs = Presentation(template) if template else Presentation()
            
            # Set slide size to widescreen (16:9)
            self.prs.slide_width = Inches(13.33)
            self.prs.slide_height = Inches(7.5)
            
            # Get common layouts
            self.layouts = {
                'title': self.prs.slide_layouts[0],
                'content': self.prs.slide_layouts[1],
                'blank': self.prs.slide_layouts[6]
            }
            
        except Exception as e:
            raise ValueError(f"Error creating PowerPoint presentation: {str(e)}. Check if PowerPoint is installed and accessible.")

    def add_logo(self, slide, logo_path, position='top-right'):
        """Add logo to slide with position options"""
        try:
            if not os.path.exists(logo_path):
                print(f"Warning: Logo file not found at {logo_path}")
                return

            # Define positions
            positions = {
                'top-right': (Inches(11.5), Inches(0.3), Inches(1.5), Inches(0.75)),
                'top-left': (Inches(0.3), Inches(0.3), Inches(1.5), Inches(0.75)),
                'bottom-right': (Inches(11.5), Inches(6.5), Inches(1.5), Inches(0.75)),
                'bottom-left': (Inches(0.3), Inches(6.5), Inches(1.5), Inches(0.75))
            }

            if position not in positions:
                position = 'top-right'

            x, y, width, height = positions[position]
            slide.shapes.add_picture(logo_path, x, y, width, height)

        except Exception as e:
            print(f"Warning: Failed to add logo to slide: {e}")

    def save_chart(self, fig, filename):
        """Save matplotlib figure as image with proper cleanup"""
        try:
            filepath = os.path.join(self.temp_dir, filename)
            fig.savefig(filepath, bbox_inches='tight', dpi=300)
            plt.close(fig)
            return filepath
        except Exception as e:
            plt.close(fig)
            raise ValueError(f"Error saving chart: {str(e)}")

    def create_title_slide(self, title, subtitle=None, logo_path=None):
        """Create title slide with optional logo"""
        slide = self.prs.slides.add_slide(self.layouts['title'])
        
        # Add title
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Add subtitle if provided
        if subtitle:
            subtitle_shape = slide.placeholders[1]
            subtitle_shape.text = subtitle
            
        # Add logo if provided
        if logo_path:
            self.add_logo(slide, logo_path)
            
        return slide

    def create_content_slide(self, title, content=None, logo_path=None):
        """Create content slide with optional logo"""
        slide = self.prs.slides.add_slide(self.layouts['content'])
        
        # Add title
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Add content if provided
        if content:
            content_shape = slide.placeholders[1]
            content_shape.text = content
            
        # Add logo if provided
        if logo_path:
            self.add_logo(slide, logo_path)
            
        return slide

    def add_chart_to_slide(self, slide, chart_path, left=Inches(1), top=Inches(2), width=Inches(8), height=Inches(4.5)):
        """Add chart image to slide with specified position and size"""
        try:
            slide.shapes.add_picture(chart_path, left, top, width, height)
        except Exception as e:
            raise ValueError(f"Error adding chart to slide: {str(e)}")

    def save(self, output_path):
        """Save presentation with proper error handling"""
        try:
            self.prs.save(output_path)
        except Exception as e:
            raise ValueError(f"Error saving presentation: {str(e)}. Check file permissions and disk space.")

def create_chart_for_ppt(df, chart_type, title, save_path):
    """Create a chart for PowerPoint slides"""
    plt.figure(figsize=(10, 6))
    
    if chart_type == 'events_by_type':
        event_counts = df['Events'].value_counts()
        plt.pie(event_counts.values, labels=event_counts.index, autopct='%1.1f%%')
    elif chart_type == 'events_by_severity':
        severity_counts = df['Severity'].value_counts()
        plt.pie(severity_counts.values, labels=severity_counts.index, autopct='%1.1f%%')
    elif chart_type == 'events_by_area':
        area_counts = df['Area'].value_counts()
        plt.bar(range(len(area_counts)), area_counts.values)
        plt.xticks(range(len(area_counts)), area_counts.index, rotation=45, ha='right')
    elif chart_type == 'weekly_trend':
        df['Date'] = pd.to_datetime(df['Date'])
        weekly_counts = df.groupby(pd.Grouper(key='Date', freq='W-MON')).size()
        plt.plot(weekly_counts.index, weekly_counts.values, marker='o')
        plt.gcf().autofmt_xdate()
        plt.grid(True)
    elif chart_type == 'monthly_trend':
        df['Date'] = pd.to_datetime(df['Date'])
        monthly_counts = df.groupby(pd.Grouper(key='Date', freq='M')).size()
        plt.plot(monthly_counts.index, monthly_counts.values, marker='o')
        plt.gcf().autofmt_xdate()
        plt.grid(True)
    
    plt.title(title)
    plt.tight_layout()
    plt.savefig(save_path)
    plt.close()

def get_accurate_kpis(df):
    """Calculate accurate KPIs from the data"""
    total_events = len(df)
    total_cameras = df['Camera'].nunique()
    total_areas = df['Area'].nunique()
    
    severity_counts = df['Severity'].value_counts()
    critical_events = severity_counts.get('Critical', 0)
    high_events = severity_counts.get('High', 0)
    
    return {
        'total_events': total_events,
        'total_cameras': total_cameras,
        'total_areas': total_areas,
        'critical_events': critical_events,
        'high_events': high_events
    }

def get_accurate_event_counts(df):
    """Get accurate event counts by type"""
    return df['Events'].value_counts()

def get_accurate_camera_data(df):
    """Get accurate camera data"""
    camera_events = df.groupby('Camera').size().sort_values(ascending=False)
    camera_areas = df.groupby('Camera')['Area'].first()
    
    return {
        'events': camera_events,
        'areas': camera_areas
    }

def get_accurate_severity_data(df):
    """Get accurate severity data"""
    severity_counts = df['Severity'].value_counts()
    severity_by_camera = df.groupby(['Camera', 'Severity']).size().unstack(fill_value=0)
    
    return {
        'counts': severity_counts,
        'by_camera': severity_by_camera
    }

def create_title_slide(prs, blank_slide_layout, customer_name=None, customer_logo_path=None, site_location=None):
    """Create customized first slide with customer information"""
    if os.path.exists('slide1.png'):
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Set white background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # Add the base template
        slide.shapes.add_picture('slide1.png', Inches(0), Inches(0), 
                               width=Inches(13.33), height=Inches(7.5))
        
        # Add customer information if provided
        if any([customer_name, site_location]):
            info_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
            text_frame = info_box.text_frame
            text_frame.word_wrap = True
            
            if customer_name:
                p = text_frame.add_paragraph()
                p.text = f"Customer: {customer_name}"
                p.font.size = Pt(24)
                p.font.bold = True
            
            if site_location:
                p = text_frame.add_paragraph()
                p.text = f"Location: {site_location}"
                p.font.size = Pt(24)
                p.font.bold = True
        
        # Add customer logo if provided
        if customer_logo_path:
            add_logo_to_slide(slide, customer_logo_path, x_position=1, y_position=0.5)
        
        # Add Visionify logo
        add_logo_to_slide(slide, 'visionify.jpeg', x_position=10.5, y_position=0.2)
        
        return slide
    return None

def create_event_trends_slide(prs, df, blank_slide_layout, custom_logo_path=None):
    """Create enhanced event trends slide with review and accuracy metrics"""
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Add title
    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9), Inches(0.5))
    title_frame = title_shape.text_frame
    title_frame.text = "Event Trends Analysis"
    title_frame.paragraphs[0].font.size = Pt(28)
    title_frame.paragraphs[0].font.bold = True
    
    # Calculate metrics
    total_events = len(df)
    review_percent = 85  # This should come from actual data
    accuracy_percent = 92  # This should come from actual data
    
    # Add metrics boxes
    metrics_data = [
        ("Total Events", total_events),
        ("Review Rate", f"{review_percent}%"),
        ("Accuracy", f"{accuracy_percent}%")
    ]
    
    for i, (label, value) in enumerate(metrics_data):
        x = 0.5 + (i * 2.5)
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1), Inches(2), Inches(0.8))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(240, 240, 240)
        box.line.color.rgb = RGBColor(200, 200, 200)
        
        text_frame = box.text_frame
        text_frame.word_wrap = True
        p = text_frame.paragraphs[0]
        p.text = f"{label}\n{value}"
        p.font.size = Pt(14)
        p.alignment = PP_ALIGN.CENTER
        p.font.bold = True
    
    # Create and add weekly trends chart
    temp_dir = tempfile.mkdtemp()
    try:
        # Weekly trends chart
        trends_path = os.path.join(temp_dir, "weekly_trends.png")
        df['Date'] = pd.to_datetime(df['Date'])
        weekly_counts = df.groupby(pd.Grouper(key='Date', freq='W-MON')).size()
        
        plt.figure(figsize=(8, 3))
        plt.plot(weekly_counts.index, weekly_counts.values, marker='o', linewidth=2)
        plt.title("Weekly Event Trends")
        plt.grid(True, alpha=0.3)
        plt.gcf().autofmt_xdate()
        
        # Calculate and add week-over-week change
        if len(weekly_counts) >= 2:
            last_week = weekly_counts.iloc[-1]
            prev_week = weekly_counts.iloc[-2]
            wow_change = ((last_week - prev_week) / prev_week) * 100
            plt.text(weekly_counts.index[-1], weekly_counts.values[-1],
                    f'\n{wow_change:+.1f}%',
                    ha='center', va='bottom')
        
        plt.tight_layout()
        plt.savefig(trends_path, bbox_inches='tight', dpi=300)
        plt.close()
        
        # Add chart to slide
        slide.shapes.add_picture(trends_path, Inches(0.5), Inches(2),
                               width=Inches(12), height=Inches(4))
        
    finally:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    # Add logo
    add_logo_to_slide(slide, custom_logo_path)
    
    return slide

def create_high_impact_slide(prs, df, blank_slide_layout, custom_logo_path=None):
    """Create enhanced high impact analysis slide with top contributing events"""
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Add title
    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9), Inches(0.5))
    title_frame = title_shape.text_frame
    title_frame.text = "High Impact Analysis"
    title_frame.paragraphs[0].font.size = Pt(28)
    title_frame.paragraphs[0].font.bold = True
    
    # Calculate high impact events
    high_impact_events = len(df[df['Severity'].isin(['Critical', 'High'])])
    
    # Get top 3 contributing event types
    top_events = df[df['Severity'].isin(['Critical', 'High'])]['Events'].value_counts().head(3)
    
    # Add high impact summary
    summary_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12), Inches(1))
    text_frame = summary_box.text_frame
    text_frame.word_wrap = True
    
    p = text_frame.add_paragraph()
    p.text = f"Total High Impact Events: {high_impact_events}"
    p.font.size = Pt(16)
    p.font.bold = True
    
    p = text_frame.add_paragraph()
    p.text = "Top Contributing Event Types:"
    p.font.size = Pt(14)
    
    for event_type, count in top_events.items():
        p = text_frame.add_paragraph()
        p.text = f"• {event_type}: {count} events"
        p.font.size = Pt(14)
        p.level = 1
    
    # Create visualization
    temp_dir = tempfile.mkdtemp()
    try:
        chart_path = os.path.join(temp_dir, "high_impact.png")
        plt.figure(figsize=(10, 4))
        
        # Create stacked bar chart of severity by date
        df['Date'] = pd.to_datetime(df['Date'])
        severity_by_date = df.pivot_table(
            index=pd.Grouper(key='Date', freq='W-MON'),
            columns='Severity',
            values='Events',
            aggfunc='count',
            fill_value=0
        )
        
        severity_by_date.plot(kind='bar', stacked=True)
        plt.title("Weekly High Impact Events")
        plt.xlabel("Week")
        plt.ylabel("Number of Events")
        plt.legend(title="Severity")
        plt.tight_layout()
        
        plt.savefig(chart_path, bbox_inches='tight', dpi=300)
        plt.close()
        
        # Add chart to slide
        slide.shapes.add_picture(chart_path, Inches(0.5), Inches(3),
                               width=Inches(12), height=Inches(4))
        
    finally:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    # Add logo
    add_logo_to_slide(slide, custom_logo_path)
    
    return slide

def create_camera_analysis_slide(prs, df, blank_slide_layout, custom_logo_path=None):
    """Create consolidated camera and event type analysis slide"""
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Add title
    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9), Inches(0.5))
    title_frame = title_shape.text_frame
    title_frame.text = "Camera and Event Analysis"
    title_frame.paragraphs[0].font.size = Pt(28)
    title_frame.paragraphs[0].font.bold = True
    
    # Calculate metrics
    camera_events = df.groupby('Camera').size().sort_values(ascending=False).head(10)
    event_types = df['Events'].value_counts().head(5)
    
    temp_dir = tempfile.mkdtemp()
    try:
        # Create horizontal bar chart for top cameras
        plt.figure(figsize=(10, 4))
        camera_events.plot(kind='barh')
        plt.title("Top 10 Cameras by Event Count")
        plt.xlabel("Number of Events")
        plt.tight_layout()
        
        camera_chart_path = os.path.join(temp_dir, "camera_events.png")
        plt.savefig(camera_chart_path, bbox_inches='tight', dpi=300)
        plt.close()
        
        # Add camera chart to slide
        slide.shapes.add_picture(camera_chart_path, Inches(0.5), Inches(1.5),
                               width=Inches(12), height=Inches(2.5))
        
        # Add event types summary as text
        text_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(12), Inches(2))
        text_frame = text_box.text_frame
        
        p = text_frame.add_paragraph()
        p.text = "Top 5 Event Types:"
        p.font.size = Pt(16)
        p.font.bold = True
        
        for event_type, count in event_types.items():
            p = text_frame.add_paragraph()
            p.text = f"• {event_type}: {count} events"
            p.font.size = Pt(14)
            p.level = 1
        
    finally:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    # Add logo
    add_logo_to_slide(slide, custom_logo_path)
    
    return slide

def create_recommendations_slide(prs, blank_slide_layout, custom_logo_path=None):
    """Create a template for recommendations slide"""
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Add title
    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9), Inches(0.5))
    title_frame = title_shape.text_frame
    title_frame.text = "Recommendations"
    title_frame.paragraphs[0].font.size = Pt(28)
    title_frame.paragraphs[0].font.bold = True
    
    # Add placeholder text boxes for recommendations
    text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12), Inches(5))
    text_frame = text_box.text_frame
    
    p = text_frame.add_paragraph()
    p.text = "[Add Key Recommendations Here]"
    p.font.size = Pt(18)
    p.font.bold = True
    
    for i in range(3):
        p = text_frame.add_paragraph()
        p.text = f"[Recommendation {i+1} Details]"
        p.font.size = Pt(14)
        p.level = 1
    
    # Add note about manual updates
    note_box = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(12), Inches(0.5))
    note_frame = note_box.text_frame
    p = note_frame.add_paragraph()
    p.text = "Note: Please update recommendations manually based on analysis and insights."
    p.font.size = Pt(10)
    p.font.italic = True
    
    # Add logo
    add_logo_to_slide(slide, custom_logo_path)
    
    return slide

def create_powerpoint_report(df_selected, report_title="Visionify Weekly Report", 
                           customer_name=None, customer_logo_path=None, site_location=None):
    """Generate a professional PowerPoint presentation with analytics and default slides"""
    # Create presentation
    prs = Presentation()
    
    # Set slide size to widescreen (16:9)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # Set up slide layouts
    blank_slide_layout = prs.slide_layouts[6]  # Blank slide
    
    # Create title slide with customer information
    create_title_slide(prs, blank_slide_layout, customer_name, customer_logo_path, site_location)
    
    # Create enhanced event trends slide
    create_event_trends_slide(prs, df_selected, blank_slide_layout, customer_logo_path)
    
    # Create high impact analysis slide
    create_high_impact_slide(prs, df_selected, blank_slide_layout, customer_logo_path)
    
    # Create consolidated camera analysis slide
    create_camera_analysis_slide(prs, df_selected, blank_slide_layout, customer_logo_path)
    
    # Create recommendations slide
    create_recommendations_slide(prs, blank_slide_layout, customer_logo_path)
    
    # Add remaining default slides (Floor Dashboard, Supported AI Scenarios, Contact)
    for filename in ['slide12.png', 'slide13.png', 'slide14.png']:
        if os.path.exists(filename):
            slide = prs.slides.add_slide(blank_slide_layout)
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 255, 255)
            slide.shapes.add_picture(filename, Inches(0), Inches(0), 
                                   width=Inches(13.33), height=Inches(7.5))
            add_logo_to_slide(slide, customer_logo_path)
    
    # Save to buffer
    ppt_buffer = BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    
    return ppt_buffer