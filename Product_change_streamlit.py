import streamlit as st
import openpyxl
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill, Font

def merge_inv_rpt(Inv_List):  
    # report workbook before being saved
    Rpt = Workbook()
    Rpt.remove(Rpt.active)
    # read all tank inventory excel files
    for Inv_file in Inv_List:
        # open each product inventory workbook
        Inv = openpyxl.load_workbook(Inv_file)
        for sheet in Inv.sheetnames:
             # open each product inventory sheet
            Inv_sheet = Inv[sheet]
             # copy each product inventory sheet to the report workbook
            Rpt_sheet = Rpt.create_sheet(Inv_file.name.replace('.xlsx', ''))
            for row in Inv_sheet.iter_rows(values_only=True):
                Rpt_sheet.append(row)
            # apply borders to the report workbook
            for row in Rpt_sheet.iter_rows(min_row=1, max_row=Rpt_sheet.max_row, min_col=1, max_col=Rpt_sheet.max_column):
                for cell in row:
                    cell.border = Border(left=Side(border_style='thin', color='FF000000'),
                                         right=Side(border_style='thin', color='FF000000'),
                                         top=Side(border_style='thin', color='FF000000'),
                                         bottom=Side(border_style='thin', color='FF000000')) 
                    # adjust column width   
                    Rpt_sheet.column_dimensions['A'].width = 15
                    Rpt_sheet.column_dimensions['B'].width = 20
                    Rpt_sheet.column_dimensions['C'].width = 20  
                    Rpt_sheet.column_dimensions['D'].width = 15  
                    Rpt_sheet.column_dimensions['E'].width = 15

    for m in range (0,6):
        Rpt.worksheets[m].delete_rows(1,4)
        Rpt.worksheets[m].delete_cols(2)
        Rpt.worksheets[m].delete_cols(4)
        Rpt.worksheets[m].delete_cols(5)
        Rpt.worksheets[m].delete_cols(6,2)
    # fill first row with yellow color and apply bold font to tank inventory tabs for each site   
        for n in ['A1','B1','C1','D1','E1']:
            Rpt.worksheets[m][n].fill = PatternFill('solid', start_color='00FFFFCC')
            Rpt.worksheets[m][n].font = Font(bold=True)
    return Rpt
#Streamlit app
st.set_page_config(layout="wide", page_title="Product Change Report")
st.write("## Product change report for GL")
st.write("Upload current PAS.xlsx, GP.xlsx, GPWC.xlsx, JSTR.xlsx, KMET.xlsx and BOSTCO.xlsx")
new_rpt = st.sidebar.text_input("Current Report Date (mm-dd-yyyy):")
ProductChange = 'Product Change '+ new_rpt + '.xlsx'

Inv_List = st.file_uploader("Choose `.xlsx` files", type="xlsx", accept_multiple_files=True)

if Inv_List:
    merged_workbook = merge_inv_rpt(Inv_List)
    merged_workbook.save(ProductChange)
    
    with open(ProductChange, "rb") as file:
        st.download_button(label="Download Merged Workbook", data=file, file_name = ProductChange, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
