# -*- coding: utf-8 -*-
"""
Created on Wed Nov 22 14:17:16 2023

@author: haalo
"""

from PyPDF2 import PdfReader
import xlsxwriter
import os
def specified_extraction(start, end, text, start_order):
    extraction = False
    extracted_text = []
    current_order = 0
    for word in text:  
        if start.lower() == word.lower():
            current_order += 1
            if current_order == start_order:
                extraction = True            
        if extraction:
            extracted_text += [word]            
        if end.lower() == word.lower():
            extraction = False       
    return extracted_text[1:-1]


try:    
    
    print("Enter path to any article for extraction (No quotations) (type 'stop' to end)")
    print("Enter last name of the author and the year for each article.")
    articles_dict = {}
    article_names = {}
    while True:
        
        pdf_path = input("Enter pdf file path (no quotations): ")
                
        if pdf_path == 'stop':
            
            directory_path = input("Enter directory path (no quotations): ")
            os.chdir(directory_path)
            
            xlfile_name = input("Enter excel file name (ending with '.xlsx'): ")
            
            workbook = xlsxwriter.Workbook(xlfile_name)
            worksheet = workbook.add_worksheet()
            
            worksheet.write('A1', 'Articles')
            worksheet.write('B1', 'Introduction')
            worksheet.write('C1', 'Methods')
            worksheet.write('D1', 'Results')            
            row = 1  
            for key, value in articles_dict.items():
                worksheet.write(row, 0, key)  
                worksheet.write(row, 1, value[0])  
                worksheet.write(row, 2, value[1])
                worksheet.write(row, 3, value[2])  
                row += 1
            
            workbook.close()
            print('File generated. Check specified directory.')
            break
        
        else:
            
            article_id = input("Author and year: ")        
            article_names[pdf_path] = article_id
            
            reader = PdfReader(pdf_path, "rb")    
            extracted_text = ""
            for i in range(0, len(reader.pages)):
                selected_page = reader.pages[i]
                text = selected_page.extract_text()
                extracted_text += text
                final_text = extracted_text.split()    
        
        introduction_list = specified_extraction('Introduction', 'Methods', final_text, 1)
        methods_list = specified_extraction('Methods', 'Results', final_text, 1)
        results_list = specified_extraction('Results', 'Discussion', final_text, 1)
        
        sign = ' '
        introduction_text = sign.join(introduction_list)
        methods_text = sign.join(methods_list)
        results_text = sign.join(results_list)
        section_list = [introduction_text, methods_text, results_text]
        articles_dict[article_names[pdf_path]] = section_list        
        
        
except FileNotFoundError:
    print("File not found. Please try again.")
except OSError:
    print("Directory path not found. Please try again.")
except Exception:
    print("Something went wrong. Please try again.")    