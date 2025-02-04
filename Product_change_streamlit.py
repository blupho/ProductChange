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
    # read all tank inventory excel files
    for Inv_file in Inv_List:
        # open each product inventory workbook
        Inv = openpyxl.load_workbook(Inv_file)
        Inv_sheet = Inv.active
        # create a new report worksheet with the name of the product inventory workbook file name    
        Rpt_sheet = Rpt.create_sheet(Inv_file.replace('.xlsx',''))
        for i in range (1, Inv_sheet.max_row + 1 ):
            for j in range (1,Inv_sheet.max_column + 1):
                # read each cell value from product inventory worksheet            
                k = Inv_sheet.cell(row = i, column = j).value
                # write the copied value into the new report worksheet
                Rpt_sheet.cell(row = i, column = j).value = k
                # add border to each cell           
                Rpt_sheet.cell(row = i, column = j).border = Border(
                                                                        left=Side(style='thin'), 
                                                                        right=Side(style='thin'), 
                                                                        top=Side(style='thin'), 
                                                                        bottom=Side(style='thin')
                                                                        )       
        # adjust column width   
        Rpt_sheet.column_dimensions['A'].width = 15
        Rpt_sheet.column_dimensions['B'].width = 20
        Rpt_sheet.column_dimensions['C'].width = 20  
        Rpt_sheet.column_dimensions['D'].width = 15  
        Rpt_sheet.column_dimensions['E'].width = 15
    # delete the extra blank tab
    del Rpt['Sheet']
    # delete extra rows and columns from the original tank inventory files
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
Inv_Rpt = st.file_uploader('Multiple Excel files', type=["xlsx"], accept_multiple_files=True)
new_rpt = st.sidebar.text_input("Current Report Date (mm-dd-yyyy):")
ProductChange = 'Product Change '+ new_rpt + '.xlsx'
if uploaded_files:
    merged_workbook = merge_inv_rpt(Inv_Rpt)
    merged_workbook.save(ProductChange)
    
    with open(ProductChange, "rb") as file:
        st.download_button(label="Download Merged Workbook", data=file, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
