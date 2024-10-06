import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

def excel_to_powerpoint(excel_file, ppt_template=None):
    # Load the Excel file
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    sheet = wb.active
    
    # Create or Load PowerPoint Presentation
    if ppt_template:
        prs = Presentation(ppt_template)
    else:
        prs = Presentation()
        
    # Add title slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Weekly Financial Report"
    subtitle.text = f"Data from: {excel_file}"
    
    # Function to add a text slide
    def add_text_slide(title, content):
        bullet_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        title_shape.text = title
        tf = body_shape.text_frame
        tf.text = content
        
    # Add summary slide
    summary = f"Total Revenue: ${sheet['B2'].value:,.2f}\n"
    summary += f"Total Expenses: ${sheet['B3'].value:,.2f}\n"
    summary += f"Net Profit: ${sheet['B4'].value:,.2f}"
    add_text_slide("Financial Summary", summary)
    
    # Add chart slide
    chart_slid_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(chart_slide_layout)
    shapes = slide.shapes
    title_shape = shapes.title
    title_shape.text = 'Revenue vs Expenses'
    
    # Add chart data 
    chart_data = CategoryChartData()
    chart_data.categories = ['Revenue', 'Expenses']
    chart_data.add_series('Amount', (sheet['B2'].value, sheet['B3'].value))
    
    # Add chart to slide 
    x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
    chart = shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y,cx, cy, chart_data
    ).chart
    
    # Save the presentation
    ppt_file = excel_file.replace('.xlsx', '.pptx')
    prs.save(ppt_file)
    print(f"PowerPoint presentation saved as {ppt_file}")
    
# Usage
excel_to_powerpoint('weekly_financial_report.xlsx')
        
