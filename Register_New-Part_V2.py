import pandas as pd
import os
from tkinter import Tk
import tkinter.filedialog as filedialog
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from datetime import datetime

# Step 01: 팝업에서 원본 엑셀파일을 선택한다.
Tk().withdraw()  # Close the root window
file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xls;*.xlsx")])
if not file_path:
    raise ValueError("No file selected. Program terminated.")

# 저장할 폴더를 설정한다.
save_folder = 'C:/신규파트등록/'
today_date = datetime.today().strftime('%Y%m%d')
save_path = os.path.join(save_folder, today_date)

# 날짜별 폴더 생성
save_folder = os.path.join(save_folder, datetime.today().strftime('%Y%m%d'))
os.makedirs(save_folder, exist_ok=True)

# Step 03: 원본 엑셀파일에서 BOX열을 찾고 값이 16이면 Material열 값을 BOX_16 배열에 저장, BOX열 값이 16이 아니면 Material열 값을 BOX_XX 배열에 저장한다.
df = pd.read_excel(file_path)
box_16 = df[df['BOX'] == 16]['Material'].tolist()
box_xx = df[df['BOX'] != 16]['Material'].tolist()

# 첫 번째 출력 파일 저장 (파일 분할 로직 추가)
max_rows = 9999
file_index = 1
for i in range(0, len(box_xx), max_rows):
    chunk = box_xx[i:i + max_rows]
    output_file = os.path.join(save_folder, f'5-1. QM_일반-{file_index}.xlsx')
    
    # Create DataFrame with required structure
    output_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'QM proc. active': [''] * len(chunk),
        'QM control key': [''] * len(chunk),
        'Insp type': ['01'] * len(chunk),
        'Active': ['x'] * len(chunk)
    })
    
    additional_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'QM proc. active': ['x'] * len(chunk),
        'QM control key': ['0001'] * len(chunk),
        'Insp type': ['0130'] * len(chunk),
        'Active': ['x'] * len(chunk)
    })
    
    final_df = pd.concat([output_df, additional_df], ignore_index=True)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        
        # Set column width
        column_widths = [12.3, 9.6, 19.8, 19.5, 13.4, 10.6]
        for col_num, width in enumerate(column_widths, start=1):
            worksheet.column_dimensions[chr(64 + col_num)].width = width
        
        # Apply text format
        for col in worksheet.columns:
            for cell in col:
                cell.alignment = Alignment(horizontal='left')
                cell.number_format = "@"
    
    file_index += 1

# 두 번째 출력 파일 저장 (파일 분할 로직 추가)
max_rows = 9999
file_index = 1
for i in range(0, len(box_xx), max_rows):
    chunk = box_xx[i:i + max_rows]
    output_file = os.path.join(save_folder, f'6-1. Insp_일반-{file_index}.xlsx')
    
    # Create DataFrame with required structure
    output_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'Date': [datetime.today().strftime('%Y%m%d')] * len(chunk),
        'Usage': [''] * len(chunk),
        'Status': ['5'] * len(chunk),
        'Dynamic mod level': ['4'] * len(chunk),
        'Modification rule': [''] * len(chunk),
        'Control key': ['qm02'] * len(chunk),
        'Description': ['수입검사'] * len(chunk),
        'Inspection': ['CHL00001'] * len(chunk),
        'Sampling': ['fo001'] * len(chunk)
    })
    
    additional_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'Date': [datetime.today().strftime('%Y%m%d')] * len(chunk),
        'Usage': ['50'] * len(chunk),
        'Status': [''] * len(chunk),
        'Dynamic mod level': ['4'] * len(chunk),
        'Modification rule': [''] * len(chunk),
        'Control key': ['qm02'] * len(chunk),
        'Description': ['출장검사'] * len(chunk),
        'Inspection': ['CHL00001'] * len(chunk),
        'Sampling': ['fo001'] * len(chunk)
    })
    
    final_df = pd.concat([output_df, additional_df], ignore_index=True)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        
        # Set column width
        column_widths = [12.3, 9.6, 9.3, 10.5, 10.6, 22.9, 21, 15.5, 15.3, 14.3, 13.3]
        for col_num, width in enumerate(column_widths, start=1):
            worksheet.column_dimensions[chr(64 + col_num)].width = width
        
        # Apply text format
        for col in worksheet.columns:
            for cell in col:
                cell.alignment = Alignment(horizontal='left')
                cell.number_format = "@"
    
    file_index += 1

# 세 번째 출력 파일 저장 (파일 분할 로직 추가)
max_rows = 9999
file_index = 1
for i in range(0, len(box_16), max_rows):
    chunk = box_16[i:i + max_rows]
    output_file = os.path.join(save_folder, f'5-2. QM_16BOX (복사 3번)-{file_index}.xlsx')
    
    # Create DataFrame with required structure
    output_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'QM proc. active': [''] * len(chunk),
        'QM control key': [''] * len(chunk),
        'Insp type': ['01'] * len(chunk),
        'Active': ['x'] * len(chunk)
    })
    
    additional_df_1 = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'QM proc. active': [''] * len(chunk),
        'QM control key': [''] * len(chunk),
        'Insp type': ['06'] * len(chunk),
        'Active': ['x'] * len(chunk)
    })
    
    additional_df_2 = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'QM proc. active': ['x'] * len(chunk),
        'QM control key': ['0001'] * len(chunk),
        'Insp type': ['0130'] * len(chunk),
        'Active': ['x'] * len(chunk)
    })
    
    final_df = pd.concat([output_df, additional_df_1, additional_df_2], ignore_index=True)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        
        # Set column width
        column_widths = [12.3, 9.6, 19.8, 19.5, 13.4, 10.6]
        for col_num, width in enumerate(column_widths, start=1):
            worksheet.column_dimensions[chr(64 + col_num)].width = width
        
        # Apply text format
        for col in worksheet.columns:
            for cell in col:
                cell.alignment = Alignment(horizontal='left')
                cell.number_format = "@"
    
    file_index += 1

# 네 번째 출력 파일 저장 (파일 분할 로직 추가)
max_rows = 9999
file_index = 1
for i in range(0, len(box_16), max_rows):
    chunk = box_16[i:i + max_rows]
    output_file = os.path.join(save_folder, f'6-2. Insp_16BOX(수입, 출장검사)-{file_index}.xlsx')
    
    # Create DataFrame with required structure
    output_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'Date': [datetime.today().strftime('%Y%m%d')] * len(chunk),
        'Usage': ['5'] * len(chunk),
        'Status': ['4'] * len(chunk),
        'Dynamic mod level': [''] * len(chunk),
        'Modification rule': [''] * len(chunk),
        'Control key': ['qm02'] * len(chunk),
        'Description': ['수입검사'] * len(chunk),
        'Inspection': ['CHL00001'] * len(chunk),
        'Sampling': ['fo001'] * len(chunk)
    })
    
    additional_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'Date': [datetime.today().strftime('%Y%m%d')] * len(chunk),
        'Usage': ['50'] * len(chunk),
        'Status': ['4'] * len(chunk),
        'Dynamic mod level': [''] * len(chunk),
        'Modification rule': [''] * len(chunk),
        'Control key': ['qm02'] * len(chunk),
        'Description': ['출장검사'] * len(chunk),
        'Inspection': ['CHL00001'] * len(chunk),
        'Sampling': ['fo001'] * len(chunk)
    })
    
    final_df = pd.concat([output_df, additional_df], ignore_index=True)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        
        # Set column width
        column_widths = [12.3, 9.6, 9.3, 10.5, 10.6, 22.9, 21, 15.5, 15.3, 14.3, 13.3]
        for col_num, width in enumerate(column_widths, start=1):
            worksheet.column_dimensions[chr(64 + col_num)].width = width
        
        # Apply text format
        for col in worksheet.columns:
            for cell in col:
                cell.alignment = Alignment(horizontal='left')
                cell.number_format = "@"
    
    file_index += 1

# 다섯 번째 출력 파일 저장 (파일 분할 로직 추가)
max_rows = 9999
file_index = 1
for i in range(0, len(box_16), max_rows):
    chunk = box_16[i:i + max_rows]
    output_file = os.path.join(save_folder, f'6-3. Insp_16BOX(수리검사) (복사 1번)-{file_index}.xlsx')
    
    # Create DataFrame with required structure
    output_df = pd.DataFrame({
        'Material': chunk,
        'Plant': ['1001'] * len(chunk),
        'Date': [datetime.today().strftime('%Y%m%d')] * len(chunk),
        'Usage': ['54'] * len(chunk),
        'Status': ['4'] * len(chunk),
        'Dynamic mod level': [''] * len(chunk),
        'Modification rule': [''] * len(chunk),
        'Control key': ['qm02'] * len(chunk),
        'Description': ['수리검사'] * len(chunk),
        'Inspection': ['CHL00003'] * len(chunk),
        'Sampling': ['to001'] * len(chunk)
    })
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        
        # Set column width
        column_widths = [12.3, 9.6, 9.3, 10.5, 10.6, 22.9, 21, 15.5, 15.3, 14.3, 13.3]
        for col_num, width in enumerate(column_widths, start=1):
            worksheet.column_dimensions[chr(64 + col_num)].width = width
        
        # Apply text format
        for col in worksheet.columns:
            for cell in col:
                cell.alignment = Alignment(horizontal='left')
                cell.number_format = "@"
    
    file_index += 1

# Notify success
print(f"Files saved successfully in folder: {save_folder}")
