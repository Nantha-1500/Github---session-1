import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import os

def create_test_summary_report():
    """
    Create a Test Summary Report based on the template format
    and store it in the outputs folder
    """
    
    # Load the template
    template_path = "templates/MSME-Base-Test-Cases-V1.xlsx"
    wb = openpyxl.load_workbook(template_path)
    
    # Create a new workbook for the summary report
    summary_wb = openpyxl.Workbook()
    summary_ws = summary_wb.active
    summary_ws.title = "Test Summary"
    
    # Define styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Add report header
    summary_ws['A1'] = "TEST SUMMARY REPORT"
    summary_ws['A1'].font = Font(bold=True, size=14)
    summary_ws.merge_cells('A1:F1')
    
    # Add report metadata
    row = 3
    summary_ws[f'A{row}'] = "Report Generated:"
    summary_ws[f'B{row}'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row += 1
    
    summary_ws[f'A{row}'] = "Template Used:"
    summary_ws[f'B{row}'] = "MSME-Base-Test-Cases-V1.xlsx"
    row += 2
    
    # Add column headers for summary table
    headers = ["Test Case ID", "Test Case Name", "Status", "Result", "Execution Date", "Remarks"]
    for col_idx, header in enumerate(headers, 1):
        cell = summary_ws.cell(row=row, column=col_idx)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    
    row += 1
    
    # Sample test data (can be populated with actual test results)
    sample_data = [
        ["TC001", "User Login", "Pass", "Successful", datetime.now().strftime("%Y-%m-%d"), "No issues"],
        ["TC002", "User Logout", "Pass", "Successful", datetime.now().strftime("%Y-%m-%d"), "No issues"],
        ["TC003", "Data Validation", "Fail", "Failed", datetime.now().strftime("%Y-%m-%d"), "Data mismatch"],
        ["TC004", "API Response", "Pass", "Successful", datetime.now().strftime("%Y-%m-%d"), "Response time: 200ms"],
    ]
    
    for data_row in sample_data:
        for col_idx, value in enumerate(data_row, 1):
            cell = summary_ws.cell(row=row, column=col_idx)
            cell.value = value
            cell.border = border
            cell.alignment = Alignment(horizontal='left', vertical='center')
        row += 1
    
    # Add summary statistics
    row += 2
    summary_ws[f'A{row}'] = "SUMMARY STATISTICS"
    summary_ws[f'A{row}'].font = Font(bold=True, size=11)
    row += 1
    
    summary_ws[f'A{row}'] = "Total Test Cases:"
    summary_ws[f'B{row}'] = len(sample_data)
    row += 1
    
    summary_ws[f'A{row}'] = "Passed:"
    summary_ws[f'B{row}'] = sum(1 for data in sample_data if data[2] == "Pass")
    row += 1
    
    summary_ws[f'A{row}'] = "Failed:"
    summary_ws[f'B{row}'] = sum(1 for data in sample_data if data[2] == "Fail")
    row += 1
    
    pass_rate = (sum(1 for data in sample_data if data[2] == "Pass") / len(sample_data)) * 100
    summary_ws[f'A{row}'] = "Pass Rate:"
    summary_ws[f'B{row}'] = f"{pass_rate:.2f}%"
    
    # Adjust column widths
    summary_ws.column_dimensions['A'].width = 15
    summary_ws.column_dimensions['B'].width = 20
    summary_ws.column_dimensions['C'].width = 12
    summary_ws.column_dimensions['D'].width = 15
    summary_ws.column_dimensions['E'].width = 15
    summary_ws.column_dimensions['F'].width = 20
    
    # Create outputs directory if it doesn't exist
    os.makedirs("outputs", exist_ok=True)
    
    # Save the report
    output_filename = f"outputs/Test_Summary_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    summary_wb.save(output_filename)
    
    print(f"✓ Test Summary Report created successfully!")
    print(f"✓ File saved to: {output_filename}")
    return output_filename

if __name__ == "__main__":
    create_test_summary_report()
