from flask import Flask, render_template, request, send_file
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime, timedelta
import os
from dateutil.relativedelta import relativedelta

app = Flask(__name__)

# Database configuration
DB_CONFIG = {
    'host': 'localhost',
    'database': 'hikvision',
    'user': 'postgres',
    'password': 'postgres',
    'port': '5432'
}

def get_db_connection():
    """Create database connection"""
    connection_string = f"postgresql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
    return create_engine(connection_string)

def generate_report(start_date, end_date, report_type):
    """Generate report based on date range"""
    engine = get_db_connection()
    
    query = """
        SELECT 
            id,
            date_and_time,
            date,
            time,
            device_name,
            reader_name,
            person_group
        FROM users
        WHERE date BETWEEN %s AND %s
        ORDER BY date_and_time
    """
    
    df = pd.read_sql_query(query, engine, params=[start_date, end_date])
    
    # Generate Excel file
    output_file = f'reports/{report_type}_report_{start_date}_{end_date}.xlsx'
    os.makedirs('reports', exist_ok=True)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write main data
        df.to_excel(writer, sheet_name='Detailed Data', index=False, startrow=1)
        
        # Add summary sheet
        summary = pd.DataFrame({
            'Total Records': [len(df)],
            'Unique Devices': [df['device_name'].nunique()],
            'Unique Groups': [df['person_group'].nunique()],
            'Date Range': [f"{start_date} to {end_date}"],
            'Report Type': [report_type.capitalize()]
        })
        summary.to_excel(writer, sheet_name='Summary', index=False)
        
        # Format the sheets
        workbook = writer.book
        for sheet_name in ['Detailed Data', 'Summary']:
            worksheet = writer.sheets[sheet_name]
            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 15
            worksheet.column_dimensions['D'].width = 15
            worksheet.column_dimensions['E'].width = 20
            worksheet.column_dimensions['F'].width = 20
            worksheet.column_dimensions['G'].width = 20
    
    return output_file

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    report_type = request.form['report_type']
    today = datetime.now().date()
    
    if report_type == 'daily':
        start_date = today
        end_date = today
    elif report_type == 'monthly':
        start_date = today.replace(day=1)
        end_date = (start_date + relativedelta(months=1, days=-1))
    elif report_type == 'quarterly':
        current_quarter = (today.month - 1) // 3
        start_date = today.replace(month=current_quarter * 3 + 1, day=1)
        end_date = (start_date + relativedelta(months=3, days=-1))
    
    output_file = generate_report(start_date, end_date, report_type)
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)