import pandas as pd
import json
from datetime import datetime
from collections import Counter
import os

def analyze_excel_file(excel_file):
    """Analyze Excel file and extract test data"""
    
    try:
        # Read Excel file
        excel_file_obj = pd.ExcelFile(excel_file)
        
        print(f"\n{'='*80}")
        print("ANALYZING EXCEL FILE")
        print(f"{'='*80}\n")
        print(f"File: {excel_file}")
        print(f"Sheet names: {excel_file_obj.sheet_names}\n")
        
        analysis_results = {}
        
        # Analyze each sheet
        for sheet in excel_file_obj.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet)
            
            print(f"\n{'─'*80}")
            print(f"Sheet: {sheet}")
            print(f"{'─'*80}")
            print(f"Total Rows: {len(df)}")
            print(f"Total Columns: {len(df.columns)}")
            print(f"\nColumns: {list(df.columns)}\n")
            print(df.head(15))
            
            analysis_results[sheet] = df
        
        return analysis_results, excel_file_obj.sheet_names
    
    except FileNotFoundError:
        print(f"Error: File '{excel_file}' not found!")
        return None, None
    except Exception as e:
        print(f"Error reading file: {str(e)}")
        return None, None

def generate_test_summary_report(data_dict):
    """Generate comprehensive test summary report"""
    
    report = {
        "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "test_summary": {},
        "overall_statistics": {}
    }
    
    total_test_cases = 0
    total_executed = 0
    total_passed = 0
    total_failed = 0
    total_skipped = 0
    
    # Analyze each sheet
    for sheet_name, df in data_dict.items():
        print(f"\n\nAnalyzing sheet: {sheet_name}")
        
        sheet_summary = {
            "total_rows": len(df),
            "columns": list(df.columns),
            "test_metrics": {}
        }
        
        # Look for test-related columns
        columns_lower = [col.lower() for col in df.columns]
        
        # Detect column names for test status
        status_col = None
        test_case_col = None
        
        for col in df.columns:
            col_lower = col.lower()
            if any(keyword in col_lower for keyword in ['status', 'result', 'outcome', 'test status']):
                status_col = col
            if any(keyword in col_lower for keyword in ['test case', 'testcase', 'tc', 'test_id']):
                test_case_col = col
        
        # Count test cases
        total_count = len(df)
        
        # Count by status if status column exists
        if status_col and df[status_col].notna().any():
            status_counts = df[status_col].value_counts()
            print(f"Status column found: {status_col}")
            print(f"Status distribution:\n{status_counts}\n")
            
            sheet_summary["test_metrics"]["total_test_cases"] = total_count
            
            passed_count = 0
            failed_count = 0
            executed_count = 0
            skipped_count = 0
            
            # Count statuses
            for status, count in status_counts.items():
                status_str = str(status).lower()
                if any(keyword in status_str for keyword in ['pass', 'passed', 'success']):
                    passed_count = count
                elif any(keyword in status_str for keyword in ['fail', 'failed', 'error']):
                    failed_count = count
                elif any(keyword in status_str for keyword in ['skip', 'skipped']):
                    skipped_count = count
                executed_count += count
            
            sheet_summary["test_metrics"]["total_executed"] = executed_count
            sheet_summary["test_metrics"]["test_passed"] = passed_count
            sheet_summary["test_metrics"]["test_failed"] = failed_count
            sheet_summary["test_metrics"]["test_skipped"] = skipped_count
            sheet_summary["test_metrics"]["pass_rate"] = round((passed_count / executed_count * 100) if executed_count > 0 else 0, 2)
            sheet_summary["test_metrics"]["fail_rate"] = round((failed_count / executed_count * 100) if executed_count > 0 else 0, 2)
            
            total_test_cases += total_count
            total_executed += executed_count
            total_passed += passed_count
            total_failed += failed_count
            total_skipped += skipped_count
        
        # Add status distribution
        if status_col:
            sheet_summary["status_distribution"] = df[status_col].value_counts().to_dict()
        
        report["test_summary"][sheet_name] = sheet_summary
    
    # Calculate overall statistics
    if total_executed > 0:
        overall_pass_rate = round((total_passed / total_executed * 100), 2)
        overall_fail_rate = round((total_failed / total_executed * 100), 2)
    else:
        overall_pass_rate = 0
        overall_fail_rate = 0
    
    report["overall_statistics"] = {
        "total_test_cases": total_test_cases,
        "total_executed": total_executed,
        "test_passed": total_passed,
        "test_failed": total_failed,
        "test_skipped": total_skipped,
        "pass_rate_percentage": overall_pass_rate,
        "fail_rate_percentage": overall_fail_rate,
        "execution_rate_percentage": round((total_executed / total_test_cases * 100) if total_test_cases > 0 else 0, 2)
    }
    
    return report

def create_markdown_report(report, output_file="TEST_SUMMARY_REPORT.md"):
    """Create a formatted Markdown report"""
    
    markdown_content = f"""# Test Summary Report

**Report Generated:** {report['report_date']}

---

## 📊 Overall Statistics

| Metric | Value |
|--------|-------|
| **Total Test Cases** | {report['overall_statistics']['total_test_cases']} |
| **Total Executed** | {report['overall_statistics']['total_executed']} |
| **Test Passed** ✅ | {report['overall_statistics']['test_passed']} |
| **Test Failed** ❌ | {report['overall_statistics']['test_failed']} |
| **Test Skipped** ⏭️ | {report['overall_statistics']['test_skipped']} |
| **Pass Rate** | {report['overall_statistics']['pass_rate_percentage']}% |
| **Fail Rate** | {report['overall_statistics']['fail_rate_percentage']}% |
| **Execution Rate** | {report['overall_statistics']['execution_rate_percentage']}% |

---

## 📈 Detailed Analysis by Sheet

"""
    
    for sheet_name, sheet_data in report['test_summary'].items():
        markdown_content += f"\n### Sheet: {sheet_name}\n\n"
        
        metrics = sheet_data['test_metrics']
        if metrics:
            markdown_content += f"""| Metric | Value |
|--------|-------|
| Total Test Cases | {metrics.get('total_test_cases', 'N/A')} |
| Total Executed | {metrics.get('total_executed', 'N/A')} |
| Passed | {metrics.get('test_passed', 'N/A')} ✅ |
| Failed | {metrics.get('test_failed', 'N/A')} ❌ |
| Skipped | {metrics.get('test_skipped', 'N/A')} ⏭️ |
| Pass Rate | {metrics.get('pass_rate', 'N/A')}% |
| Fail Rate | {metrics.get('fail_rate', 'N/A')}% |

"""
        
        if 'status_distribution' in sheet_data:
            markdown_content += "\n#### Status Distribution\n\n"
            for status, count in sheet_data['status_distribution'].items():
                markdown_content += f"- **{status}**: {count}\n"
            markdown_content += "\n"
    
    # Write to file
    with open(output_file, 'w') as f:
        f.write(markdown_content)
    
    print(f"\n✅ Markdown report saved: {output_file}")
    return markdown_content

def create_html_report(report, output_file="TEST_SUMMARY_REPORT.html"):
    """Create an HTML formatted report"""
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Test Summary Report</title>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: #f5f5f5;
                margin: 20px;
                padding: 20px;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background-color: white;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            h1 {{
                color: #333;
                border-bottom: 3px solid #007bff;
                padding-bottom: 10px;
            }}
            h2 {{
                color: #555;
                margin-top: 30px;
            }}
            .report-date {{
                color: #666;
                font-size: 14px;
                margin-bottom: 20px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
            }}
            th {{
                background-color: #007bff;
                color: white;
                padding: 12px;
                text-align: left;
            }}
            td {{
                padding: 10px;
                border-bottom: 1px solid #ddd;
            }}
            tr:hover {{
                background-color: #f9f9f9;
            }}
            .metric-box {{
                display: inline-block;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 20px;
                margin: 10px;
                border-radius: 8px;
                text-align: center;
                min-width: 150px;
            }}
            .passed {{
                background: linear-gradient(135deg, #84fab0 0%, #8fd3f4 100%);
            }}
            .failed {{
                background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
            }}
            .skipped {{
                background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
            }}
            .metric-value {{
                font-size: 24px;
                font-weight: bold;
                margin: 10px 0;
            }}
            .metric-label {{
                font-size: 14px;
                opacity: 0.9;
            }}
            .status-dist {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                gap: 15px;
                margin: 20px 0;
            }}
            .status-item {{
                background-color: #f9f9f9;
                padding: 15px;
                border-left: 4px solid #007bff;
                border-radius: 4px;
            }}
            .footer {{
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #ddd;
                text-align: center;
                color: #666;
                font-size: 12px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📋 Test Summary Report</h1>
            <p class="report-date">Report Generated: {report['report_date']}</p>
            
            <h2>📊 Overall Statistics</h2>
            <div class="metric-box">
                <div class="metric-label">Total Test Cases</div>
                <div class="metric-value">{report['overall_statistics']['total_test_cases']}</div>
            </div>
            <div class="metric-box">
                <div class="metric-label">Total Executed</div>
                <div class="metric-value">{report['overall_statistics']['total_executed']}</div>
            </div>
            <div class="metric-box passed">
                <div class="metric-label">✅ Passed</div>
                <div class="metric-value">{report['overall_statistics']['test_passed']}</div>
            </div>
            <div class="metric-box failed">
                <div class="metric-label">❌ Failed</div>
                <div class="metric-value">{report['overall_statistics']['test_failed']}</div>
            </div>
            <div class="metric-box skipped">
                <div class="metric-label">⏭️ Skipped</div>
                <div class="metric-value">{report['overall_statistics']['test_skipped']}</div>
            </div>
            
            <table>
                <tr>
                    <th>Metric</th>
                    <th>Value</th>
                </tr>
                <tr>
                    <td><strong>Pass Rate</strong></td>
                    <td><strong style="color: green;">{report['overall_statistics']['pass_rate_percentage']}%</strong></td>
                </tr>
                <tr>
                    <td><strong>Fail Rate</strong></td>
                    <td><strong style="color: red;">{report['overall_statistics']['fail_rate_percentage']}%</strong></td>
                </tr>
                <tr>
                    <td><strong>Execution Rate</strong></td>
                    <td><strong style="color: blue;">{report['overall_statistics']['execution_rate_percentage']}%</strong></td>
                </tr>
            </table>
            
            <h2>📈 Detailed Analysis by Sheet</h2>
"""
    
    for sheet_name, sheet_data in report['test_summary'].items():
        html_content += f"""
            <h3>{sheet_name}</h3>
            <table>
                <tr>
                    <th>Metric</th>
                    <th>Value</th>
                </tr>
"""
        
        metrics = sheet_data['test_metrics']
        if metrics:
            html_content += f"""
                <tr>
                    <td>Total Test Cases</td>
                    <td>{metrics.get('total_test_cases', 'N/A')}</td>
                </tr>
                <tr>
                    <td>Total Executed</td>
                    <td>{metrics.get('total_executed', 'N/A')}</td>
                </tr>
                <tr>
                    <td>✅ Passed</td>
                    <td><span style="color: green; font-weight: bold;">{metrics.get('test_passed', 'N/A')}</span></td>
                </tr>
                <tr>
                    <td>❌ Failed</td>
                    <td><span style="color: red; font-weight: bold;">{metrics.get('test_failed', 'N/A')}</span></td>
                </tr>
                <tr>
                    <td>⏭️ Skipped</td>
                    <td><span style="color: orange; font-weight: bold;">{metrics.get('test_skipped', 'N/A')}</span></td>
                </tr>
                <tr>
                    <td>Pass Rate</td>
                    <td><strong style="color: green;">{metrics.get('pass_rate', 'N/A')}%</strong></td>
                </tr>
                <tr>
                    <td>Fail Rate</td>
                    <td><strong style="color: red;">{metrics.get('fail_rate', 'N/A')}%</strong></td>
                </tr>
            </table>
"""
        
        if 'status_distribution' in sheet_data:
            html_content += "<h4>Status Distribution</h4><div class='status-dist'>"
            for status, count in sheet_data['status_distribution'].items():
                html_content += f'<div class="status-item"><strong>{status}:</strong> {count}</div>'
            html_content += "</div>"
    
    html_content += """
            <div class="footer">
                <p>Generated by Test Summary Report Generator</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    with open(output_file, 'w') as f:
        f.write(html_content)
    
    print(f"✅ HTML report saved: {output_file}")

def create_json_report(report, output_file="TEST_SUMMARY_REPORT.json"):
    """Save report as JSON"""
    
    with open(output_file, 'w') as f:
        json.dump(report, f, indent=2, default=str)
    
    print(f"✅ JSON report saved: {output_file}")

def main():
    excel_file = 'MSME-Base-Test-Cases-V1.xlsx'
    
    print("\n" + "="*80)
    print("TEST SUMMARY REPORT GENERATOR")
    print("="*80)
    
    # Analyze Excel file
    data_dict, sheet_names = analyze_excel_file(excel_file)
    
    if data_dict is None:
        return
    
    # Generate test summary report
    print("\n" + "="*80)
    print("GENERATING TEST SUMMARY REPORT")
    print("="*80)
    
    report = generate_test_summary_report(data_dict)
    
    # Display overall statistics
    print("\n" + "="*80)
    print("OVERALL TEST STATISTICS")
    print("="*80)
    
    stats = report['overall_statistics']
    print(f"""
    Total Test Cases:    {stats['total_test_cases']}
    Total Executed:      {stats['total_executed']}
    Passed Tests:        {stats['test_passed']} ✅
    Failed Tests:        {stats['test_failed']} ❌
    Skipped Tests:       {stats['test_skipped']} ⏭️
    
    Pass Rate:           {stats['pass_rate_percentage']}%
    Fail Rate:           {stats['fail_rate_percentage']}%
    Execution Rate:      {stats['execution_rate_percentage']}%
    """)
    
    # Generate reports
    print("\n" + "="*80)
    print("CREATING REPORTS")
    print("="*80 + "\n")
    
    create_markdown_report(report)
    create_html_report(report)
    create_json_report(report)
    
    print("\n" + "="*80)
    print("✅ ALL REPORTS GENERATED SUCCESSFULLY!")
    print("="*80)
    print(f"""
    Generated Files:
    1. TEST_SUMMARY_REPORT.md   (Markdown format)
    2. TEST_SUMMARY_REPORT.html (HTML format - Open in browser)
    3. TEST_SUMMARY_REPORT.json (JSON format)
    """)

if __name__ == "__main__":
    main()
