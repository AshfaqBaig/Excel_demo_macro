
# import module
import streamlit as st
import win32com.client
import pandas as pd
import os
 
# Header
st.header("Demo to run macro file from UI")

# Text Input
name = st.text_input("Enter Your name")
y1 = st.text_input("Enter Your Sale 1")
y2 = st.text_input("Enter Your Sale 2")

#instantiate excel app
xl = win32com.client.Dispatch("Excel.Application")

def run_macro_file(xl, name, sale1, sale2, PATH_TO_PDF):
    
    wb = xl.Workbooks.Open(r'macro_testing_v0.1.xlsm', ReadOnly = 1)
    sheet = wb.Worksheets('sheet1') 
    sheet.Cells(2,1).Value = name
    sheet.Cells(2,2).Value = sale1
    sheet.Cells(2,3).Value = sale2
    #MacroEnabled
    xl.Application.Run('macro_testing_v0.1.xlsm!Module1.MacroEnabled')
    result = sheet.Cells(2,4).Value 
    wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    #save the excel
    wb.Save()
    wb.Close(True)
    #xl.Application.Quit()
    return result 
    

def reset_the_excel():
    xl = win32com.client.Dispatch("Excel.Application")  #instantiate excel app
    # path =  os.getcwd().replace('\'','\\') + '\\'
    # wb = xl.Workbooks.Open(path+"macro_testing_v0.1.xlsm", ReadOnly = 1)
    wb = xl.Workbooks.Open(r'C:\Users\91997\Desktop\demo\macro_testing_v0.1.xlsm', ReadOnly = 1)
    sheet = wb.Worksheets('sheet1') 
    xl.Application.Run('macro_testing_v0.1.xlsm!Module2.Macro3')
    wb.Save()
    xl.Application.Quit()

# display the name when the submit button is clicked
# .title() is used to get the input text string
if(st.button('Submit')):
    result = name.title()
    # update_the_excel(name.title(), y1.title(), y2.title())
    PATH_TO_PDF = r'macro_testing_v0.1.pdf'
    
    result = run_macro_file(xl, name.title(), y1.title(), y2.title(), PATH_TO_PDF)
    result_str = f"Average Sales  =  {result}"
    st.success(result_str)
    st.success("Sucessfully Updated Excel")
    
# if(st.button('Reset')):
#     reset_the_excel()
#     st.success("Reset")
    
