import pandas as pd
import os
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.platypus import SimpleDocTemplate,Image, Table, TableStyle, Paragraph,Spacer,Frame
from reportlab.lib import colors,styles
from reportlab.lib.units import inch
from num2words import num2words


def load_excel_sheet(file_path, sheet_name):
    # Load the specified sheet into a DataFrameSC
    return pd.read_excel(file_path, sheet_name=sheet_name)

def get_column_names(dataframe):
    # Return a list of column names
    return dataframe.columns.tolist()

def find_row_by_cell_data(dataframe, column_name, cell_data):
    # Check if the column name exists in the DataFrame
    if column_name in dataframe.columns:
        # Find the row where the cell data matches
        matching_rows = dataframe[dataframe[column_name] == cell_data]
        if not matching_rows.empty:
            # Convert the matching row to a dictionary
            return matching_rows.iloc[0].to_dict()
        else:
            return None  # Return None if no matching data found
    else:
        return None  # Return None if the column doesn't exist

# Path to your Excel file

file_path = '/Users/niyaznabi/Desktop/Gopipay_2.xlsx'  # Update with your file path
# Name of the sheets you want to work with
sheet_name_code1 = 'Sheet1'
sheet_name_code2 = 'Sheet3'

# Load the sheets
df_code1 = load_excel_sheet(file_path, sheet_name_code1)

# Extract data from Code 2
workbook = openpyxl.load_workbook(file_path)
sheet_code2 = workbook[sheet_name_code2]

# Ask for the column name and cell data in Code 1
column_name_code1 = "Emp. Code"
cell_data_code1 = input("Enter Cell Data in Code 1: ")

# Find and display the row corresponding to the cell data in Code 1
data = find_row_by_cell_data(df_code1, column_name_code1, cell_data_code1)

# Print the extracted data from Code 1
print("\nData Extracted from Code 1:")
print(data)
print(type(data['A/C']))

# Define the create_payslip_pdf function to generate the PDF



def create_payslip(filename, data):
    doc = SimpleDocTemplate(filename, pagesize=letter, topMargin=0.5 * inch)

    styles = getSampleStyleSheet()
    styles = getSampleStyleSheet()
    Story=[]
    styles['Title'].wordWrap = 'LTR'  # Ensure that text is not wrapped
    styles['Title'].fontSize = 15  # Adjust the font size if necessary
    styles['Title'].alignment = 1  # 0=Left, 1=Center, 2=Right
    styles['Title'].leading = 14
    styles['Normal'].fontSize=10
    styles['Normal'].alignment=1
    styles['Normal'].leading = 12

    # Replace with the path to your logo image
    logo = "/Users/niyaznabi/Desktop/logo.png"
    logo_img = Image(logo, width=0.75 * inch, height=0.75 * inch)
    logo_img.hAlign = 'LEFT'

    # Replace with your company details and address
    company_name = Paragraph("SRI CHANDANA ENTERPRISES", styles['Title'])
    company_address = Paragraph("HOUSE NO. 58-1-419/3/9, MIG-77, MERRIPALEM VUDA LAYOUT, VISAKHAPATNAM - 530009",
                                styles['Normal'])

    # Create a table to hold the logo and company details side by side
    company_info_table = Table([
        [logo_img, company_name],
        ['', company_address]# The first cell is empty to align the address below the company name
    ], colWidths=[1.5 * inch, 4.5 * inch])

    company_info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (1, 1), 'TOP'),
        ('LEFTPADDING', (1, 0), (1, 0), 0),
        ('RIGHTPADDING', (1, 0), (1, 0), 0),
        ('TOPPADDING', (1, 0), (1, 0), 0),
        ('BOTTOMPADDING', (1, 0), (1, 0), 0),
        ("TOPPADDING", (1, 1), (1, 1),-40),
        ("BOTTOMPADDING", (1, 1), (1, 1), 10),
        # ('SPAN', (0, 0), (0, 0)),
        ('SPAN', (0, 0), (0, 0)),
        ('SPAN', (1, 1), (-1, -1)),  # Span the address cell across to align it below the company name
        ('SPAN', (1, 2), (-1, 2))
        # Span the address cell across to align it below the company name
    ]))

    Story.append(company_info_table)
    styles.add(ParagraphStyle(name='MonthsStyle', parent=styles['Normal'], fontSize=13,
                              spaceAfter=14))  # Adjust spaceAfter as needed

    months =Paragraph("Payslip for the month of November 2023",
                                styles['MonthsStyle'])
    styles['Normal'].fontSize=13
    print("hello",months)
    Story.append(months)

    # Employee Information Table
    employee_info = [
        ['Name:', data['Name of Workman'], 'Employee No:', data['Emp. Code']],
        ['Designation:', data['Designation'], 'Bank Name:', data['Bank']],
        ['Department:', 'SLURRY PIPELINE', 'Bank Account No.:',int( data['A/C'])],
        ['PAN No.:', data['PAN']],
        ['Effective Work Days:', data['Working Days'],'net salary:',data['Net Salary']]

    ]

    # Configure style and word wrap
    employee_info_table_style = TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ])

    employee_info_table = Table(employee_info)
    employee_info_table.setStyle(employee_info_table_style)

    Story.append(employee_info_table)
    Story.append(Spacer(1, 12))

    # Earnings and Deductions Table
    net_salry=num2words(data['Net Salary'])
    print(net_salry)
    print(data['Net Salary'])
    earnings_deductions_data = [
        ['Income','','Deductions'],
        ['Particulars','Amount','Particulars','Amount'],
        ['Basic:', data['Basic Salary'], 'PF TAX', data['P.Tax']],
        ['HRA', data['HRA'], 'ESCI NO :', data['ESIC No'] if data.get('ESIC No') else '-'],
        ['Special Allowances:',data['Spl.\n Allow'],'PF Contri:',data['Employers \n PF Contribution@13%']],
        ['Statutory Bonus:',data['Statutory Bonus'],'PF Deduction:',data['Employees \n PF Deduction @ 12%']],
        ['Earned Leaves:',data['EL']],
        ['PAN No.:', data['PAN'],'ESCI :',data['Employees ESIC Deduction @ 0.75%'] if data.get('Employees ESIC Deduction @ 0.75%') else '-'],
        ['Total:',data['Earned Salary'],'Earned CTC:',data['Earned \n CTC']],
        # ['Net Salary:', data['Net Salary'], f"{net_salry}"],
        ['Net Salary:',data['Net Salary'],f"{net_salry}"],
        ]


    earnings_deductions_table = Table(earnings_deductions_data,
                                      colWidths=[2 * inch, 1.15 * inch, 2 * inch, 1.15 * inch])
    earnings_deductions_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (2, 0), colors.lightgrey),
        ('BACKGROUND', (3, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, 1), (2, 1), colors.lightgrey),
        ('BACKGROUND', (3, 1), (-1, 1), colors.lightgrey),
        ('SPAN', (2, -1), (-1, -1)),
    ]))

    Story.append(earnings_deductions_table)
    Story.append(Spacer(1, 12))

    # System generated note
    Story.append(Paragraph("* This is a system generated payslip and does not require a signature", styles['Italic']))
    signature1_path = "/Users/niyaznabi/Desktop/sign.png"  # Replace with your signature image path # Replace with your signature image path
    #
    # # Signature images
    signature1 = Image(signature1_path, width=2 * inch, height=0.5 * inch)


    # Signature text (if you don't have signature images)
    signature_text1 = Paragraph("Authorized Signature", styles['Normal'])


    # Adding images or text to the document
    signature_table = Table([
        [signature1],  # Replace with signature_text1 and signature_text2 if using text
        [signature_text1]  # This row is for descriptive text under the signatures
    ], colWidths=[3 * inch, 3 * inch], spaceBefore=0.5 * inch)

    # Style for the signature table (optional)
    signature_table.hAlign="LEFT"
    signature_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('RIGHTPADDING',(0,0),(0,-1),60),

        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))

    # Add the signature table to the story
    Story.append(signature_table)



    doc.build(Story)

salary_slip= [data['Emp. Code'],data['Name of Workman']]
print(salary_slip)
path1="/Users/niyaznabi/Desktop/Payslip"

filename=os.path.join(path1,f"{data['Emp. Code']}_December.pdf")
print(filename)
create_payslip(filename, data)