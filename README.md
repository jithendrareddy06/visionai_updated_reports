# IVLD Safety Event Analysis Dashboard

A comprehensive Streamlit application for analyzing safety event data with PowerPoint report generation capabilities.

## 🚀 Features

### Dashboard Analytics
- **Custom Date Range Filtering**: Select specific date ranges for focused analysis
- **Data Overview**: Key metrics and event summaries
- **Interactive Charts**: Events by type, area, severity with professional visualizations
- **Trend Analysis**: Weekly and monthly trend charts
- **Heatmap Tables**: Top cameras and event types with color-coded displays
- **Area-wise Reports**: Detailed analysis for specific areas/cameras
- **Excel Comparison**: Compare multiple Excel files with percentage change analysis

### 📊 Professional PowerPoint Report Generation
Generate industry-standard, executive-ready PowerPoint presentations with comprehensive analytics:

#### 🎯 Professional Slide Structure (10 slides):
1. **Visionify AI Monitoring Solution**: Professional opening slide showcasing the solution
2. **Event Trends Dashboard**: KPI metrics and event distribution analysis
3. **High Impact Analysis**: Critical event analysis with top contributing cameras
4. **Top 10 Cameras By Event Count**: Professional heatmap table with event distribution
5. **Top 5 Event Types**: Detailed event type analysis with camera correlations
6. **Recommendations**: Professional template with structured recommendation sections
7. **Events Analysis**: Side-by-side analysis of events by type and area
8. **Safety Monitoring Overview**: Advanced monitoring capabilities and insights
9. **Analytics Dashboard**: Comprehensive data visualization and trends
10. **Contact Information**: Key contacts and next steps

#### 💼 Executive-Ready Features:
- **Professional corporate theme** with consistent branding
- **High-resolution charts (300 DPI)** for presentation quality
- **Comprehensive analytics** with area-wise breakdown
- **Data-driven insights** for stakeholder discussions
- **Structured recommendations** with priority-coded sections
- **Client-ready output** suitable for board and executive presentations
- **Excel comparison charts** showing trends and percentage changes
- **Clean chart titles** for professional presentation

## 📋 Requirements

Install dependencies using:
```bash
pip install -r requirements.txt
```

### Dependencies:
- `streamlit>=1.28.0` - Web application framework
- `pandas>=1.5.0` - Data manipulation and analysis
- `numpy>=1.21.0` - Numerical computing
- `matplotlib>=3.5.0` - Chart generation
- `openpyxl>=3.0.10` - Excel file reading
- `python-pptx>=0.6.21` - PowerPoint generation
- `Pillow>=9.0.0` - Image processing

## 🚀 Usage

### Running the Application
```bash
streamlit run app.py
```

### Data Requirements
Your Excel file must contain these columns:
- **Date** (required): Event dates (DD/MM/YYYY format)
- **Events** (required): Event type descriptions
- **Time** (optional): Event times for interval analysis
- **Severity** (optional): Event severity levels
- **Area** or **Camera** (optional): Location/area information

### Generating Professional PowerPoint Reports

1. **Upload Data**: Use the file uploader to load your Excel file
2. **Set Date Range**: Choose the analysis period using the date selectors
3. **Generate Professional Report**: 
   - Scroll to the "PowerPoint Report Generation" section
   - Customize the report title (default: "Visionify Weekly Report")
   - Click "Generate PowerPoint Report"
4. **Download**: Use the download button to save your PPTX file
5. **Review**: The presentation includes 10 professionally designed slides
6. **Present**: Use the comprehensive slide structure for effective presentations

### Excel Comparison Feature

1. **Upload Multiple Files**: Use both file uploaders for comparison
2. **Automatic Analysis**: The system calculates percentage changes between datasets
3. **Comparison Charts**: View "Average Fluctuation in Events" charts
4. **Metrics Display**: See detailed comparison metrics in the dashboard

### PowerPoint Customization Tips

- **Default Slides**: The presentation includes pre-designed company slides:
  - Slide 1: Visionify AI Monitoring Solution
  - Slide 8: Safety Monitoring Overview
  - Slide 9: Analytics Dashboard
  - Slide 10: Contact Information
- **Analytics Slides**: Slides 2-7 are automatically generated based on your data
- **Recommendations Slide**: Slide 6 provides a professional template with structured sections for manual input
- **Chart Quality**: All charts are generated at 300 DPI for professional quality
- **Layout**: Optimized for 16:9 widescreen format with consistent styling

## 🔧 Technical Features

### Data Processing
- **Smart date parsing** with error handling
- **Flexible column detection** (Area vs Camera)
- **Memory-safe chart generation** using temporary files
- **Robust error handling** with user-friendly messages
- **Excel comparison algorithms** with percentage calculations

### Chart Generation
- **High DPI output** (300 DPI) for crisp presentation quality
- **Color-coded visualizations** for easy interpretation
- **Professional styling** with grid lines and clean aesthetics
- **Automatic label wrapping** for long text
- **Clean chart titles** without dynamic percentages for clarity

### PowerPoint Integration
- **Template-free generation** using python-pptx
- **Consistent formatting** across all slides
- **Memory-efficient processing** with automatic cleanup
- **Cross-platform compatibility**
- **Structured recommendation templates** with color-coded priority sections

## 📊 Sample Output

The generated PowerPoint includes:
- Executive summary with key metrics
- Professional charts and visualizations
- Risk analysis and trend insights
- Structured recommendations template
- Excel comparison analysis (when applicable)
- Ready-to-present format

Perfect for:
- Safety committee meetings
- Executive presentations
- Compliance reporting
- Stakeholder updates
- Board presentations

## 💡 Tips for Best Results

1. **Data Quality**: Ensure consistent date formats and clean event descriptions
2. **Date Ranges**: Use meaningful date ranges for trend analysis
3. **Report Titles**: Customize titles for specific audiences or time periods
4. **Excel Comparison**: Upload both old and new datasets for trend analysis
5. **Recommendations**: Use the structured template in Slide 6 for actionable insights
6. **Review**: Always review the generated slides before presenting

## 🆘 Troubleshooting

**Common Issues:**
- **Import Errors**: Install all required dependencies from requirements.txt
- **Date Parsing**: Ensure dates are in DD/MM/YYYY format
- **Empty Charts**: Check that your data contains the required columns
- **PowerPoint Errors**: Verify python-pptx installation and permissions
- **Memory Issues**: Ensure sufficient system memory for chart generation

For technical support or feature requests, check the application logs and error messages for specific guidance.

## 🔄 Recent Updates

- **Slide Structure**: Updated to 10 professional slides with new company slides
- **Recommendations**: Enhanced Slide 6 with structured recommendation template
- **Chart Titles**: Cleaned up chart titles for better presentation clarity
- **Excel Comparison**: Added comprehensive comparison functionality
- **UI Improvements**: Streamlined interface for better user experience

---

*This dashboard provides comprehensive safety event analysis with professional PowerPoint generation capabilities, perfect for executive reporting and stakeholder communication.* 