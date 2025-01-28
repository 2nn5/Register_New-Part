# 원본 엑셀파일에서 파트번호 추출 (원본파일은 시작시 팝업창에서 선택)
# 추출된 파트번호를 정해진 횟수 반복
# 추출된 파트번호 우측에 정해진 값 입력 후 저장
# Copilot 으로 코드 작성 완료
import pandas as pd
import os
from tkinter import Tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime

# Step 01: 팝업에서 원본 엑셀파일을 선택한다.
Tk().withdraw()  # Close the root window
file_path = askopenfilename(title="신규 파트 리스트 엑셀파일을 선택해 주세요.", filetypes=[("Excel files", "*.xls;*.xlsx")])
if not file_path:
    raise ValueError("No file selected. Program terminated.")

# Step 02: 원본 엑셀파일에서 BOX열을 찾고 값이 16이면 Material열 값을 BOX_16 배열에 저장, BOX열 값이 16이 아니면 Material열 값을 BOX_XX 배열에 저장한다.
df = pd.read_excel(file_path)
box_16 = df[df['BOX'] == 16]['Material'].tolist()
box_xx = df[df['BOX'] != 16]['Material'].tolist()

# 저장할 폴더를 설정한다.
save_folder = 'D:/신규파트등록/'
# 저장할 폴더를 매번 팝업에서 선택한다.
# save_folder = filedialog.askdirectory(title="Select Folder to Save")
# if not save_folder:
#     save_folder = os.getenv("LAST_USED_FOLDER", os.path.expanduser("~"))  # Use last used folder or default to home directory
# today_date = datetime.today().strftime('%Y%m%d')
# save_path = os.path.join(save_folder, today_date)

# 디렉토리가 존재하지 않으면 생성
os.makedirs(save_path, exist_ok=True)

# Step 03-01: 첫번째 output 파일 이름을 "5-1. QM_일반.xlsx"라고 지정한다.
output_file_1 = f'{save_path}/5-1. QM_일반.xlsx'

# Step 03-02: 첫번째 output 파일 A열은 헤더를 "Material"로 입력하고, BOX_XX 배열 내 값을 가져온다.
data_1 = {
    'Material': box_xx,
    'Plant': ['1001'] * len(box_xx),
    'QM proc. active': [''] * len(box_xx),
    'QM control key': [''] * len(box_xx),
    'Insp type': ['01'] * len(box_xx),
    'Active': ['x'] * len(box_xx)
}

# Step 03-11: 아래에 다시 A열에 BOX_XX 배열 내 값을 가져온다.
# Step 03-12: B열에 "1001"값을 텍스트 형식으로 모두 넣는다.
# Step 03-13: C열에 값은 "x"값을 텍스트 형식으로 모두 넣는다.
# Step 03-14: D열에 값은 "0001"값을 텍스트 형식으로 모두 넣는다.
# Step 03-15: E열에 값은 "0130"값을 텍스트 형식으로 모두 넣는다.
# Step 03-16: F열에 값은 "x"값을 텍스트 형식으로 모두 넣는다.
additional_data_1 = {
    'Material': box_xx,
    'Plant': ['1001'] * len(box_xx),
    'QM proc. active': ['x'] * len(box_xx),
    'QM control key': ['0001'] * len(box_xx),
    'Insp type': ['0130'] * len(box_xx),
    'Active': ['x'] * len(box_xx)
}

# Combine both data sets for the first output file
combined_data_1 = {key: data_1[key] + additional_data_1[key] for key in data_1}

# Create DataFrame for the first output file
output_df_1 = pd.DataFrame(combined_data_1)

# Save the first output file to Excel
with pd.ExcelWriter(output_file_1, engine='openpyxl') as writer:
    output_df_1.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    
    for col in worksheet.columns:
        for cell in col:
            cell.number_format = "@"
    
    # Apply filter to the table
    tab = Table(displayName="Table1", ref=f"A1:F{len(output_df_1) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    worksheet.add_table(tab)

    # Set column width from A column
    worksheet.column_dimensions['A'].width = 12.3
    worksheet.column_dimensions['B'].width = 9.6
    worksheet.column_dimensions['C'].width = 19.8
    worksheet.column_dimensions['D'].width = 19.5
    worksheet.column_dimensions['E'].width = 13.4
    worksheet.column_dimensions['F'].width = 10.6

# Step 03-01: 두번째 output 파일 이름을 "6-1. Insp_일반.xlsx"라고 지정한다.
output_file_2 = f'{save_path}/6-1. Insp_일반.xlsx'

# Step 04-01: 두번째 output 파일 A열은 헤더를 "Material"로 입력하고, BOX_XX 배열 내 값을 가져온다.
data_2 = {
    'Material': box_xx,
    'Plant': ['1001'] * len(box_xx),
    'Date': [today_date] * len(box_xx),
    'Usage': [''] * len(box_xx),
    'Status': ['5'] * len(box_xx),
    'Dynamic mod level': ['4'] * len(box_xx),
    'Modification rule': [''] * len(box_xx),
    'Control key': ['qm02'] * len(box_xx),
    'Description': ['수입검사'] * len(box_xx),
    'Inspection': ['CHL00001'] * len(box_xx),
    'Sampling': ['fo001'] * len(box_xx)
}

# Step 05-01: 아래에 다시 A열에 BOX_XX 배열 내 값을 가져온다.
# Step 05-02: B열에 "1001"값을 텍스트 형식으로 모두 넣는다.
# Step 05-03: C열에 값은 오늘날짜를 "YYYYMMDD" 형식으로 모두 넣는다.
# Step 05-04: D열에 값은 "50"값을 텍스트 형식으로 모두 넣는다.
# Step 05-05: E열에 값은 입력하지 않는다.
# Step 05-06: F열에 값은 "4"값을 텍스트 형식으로 모두 넣는다.
# Step 05-07: G열에 값은 입력하지 않는다.
# Step 05-08: H열에 값은 "qm02"값을 텍스트 형식으로 모두 넣는다.
# Step 05-09: I열에 값은 "출장검사"값을 텍스트 형식으로 모두 넣는다.
# Step 05-10: J열에 값은 "CHL00001"값을 텍스트 형식으로 모두 넣는다.
# Step 05-11: K열에 값은 "fo001"값을 텍스트 형식으로 모두 넣는다.
additional_data_2 = {
    'Material': box_xx,
    'Plant': ['1001'] * len(box_xx),
    'Date': [today_date] * len(box_xx),
    'Usage': ['50'] * len(box_xx),
    'Status': [''] * len(box_xx),
    'Dynamic mod level': ['4'] * len(box_xx),
    'Modification rule': [''] * len(box_xx),
    'Control key': ['qm02'] * len(box_xx),
    'Description': ['출장검사'] * len(box_xx),
    'Inspection': ['CHL00001'] * len(box_xx),
    'Sampling': ['fo001'] * len(box_xx)
}

# Combine both data sets for the second output file
combined_data_2 = {key: data_2[key] + additional_data_2[key] for key in data_2}

# Create DataFrame for the second output file
output_df_2 = pd.DataFrame(combined_data_2)

# Save the second output file to Excel
with pd.ExcelWriter(output_file_2, engine='openpyxl') as writer:
    output_df_2.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    
    for col in worksheet.columns:
        for cell in col:
            cell.number_format = "@"
    
    # Apply filter to the table
    tab = Table(displayName="Table2", ref=f"A1:K{len(output_df_2) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    worksheet.add_table(tab)

    # Set column width from A column
    worksheet.column_dimensions['A'].width = 12.3
    worksheet.column_dimensions['B'].width = 9.6
    worksheet.column_dimensions['C'].width = 9.3
    worksheet.column_dimensions['D'].width = 10.5
    worksheet.column_dimensions['E'].width = 10.6
    worksheet.column_dimensions['F'].width = 22.9
    worksheet.column_dimensions['G'].width = 21
    worksheet.column_dimensions['H'].width = 15.5
    worksheet.column_dimensions['I'].width = 15.3
    worksheet.column_dimensions['J'].width = 14.3
    worksheet.column_dimensions['K'].width = 13.3
    
# Step 03-01: 세번째 output 파일 이름을 "5-2. QM_16BOX(복사 3번).xlsx"라고 지정한다.
output_file_3 = f'{save_path}/5-2. QM_16BOX(복사 3번).xlsx'

# Step 03-02: 세번째 output 파일 A열은 헤더를 "Material"로 입력하고, BOX_16 배열 내 값을 가져온다.
data_3 = {
    'Material': box_16,
    'Plant': ['1001'] * len(box_16),
    'QM proc. active': [''] * len(box_16),
    'QM control key': [''] * len(box_16),
    'Insp type': ['01'] * len(box_16),
    'Active': ['x'] * len(box_16)
}

# Step 04-01: 아래에 다시 A열에 BOX_16 배열 내 값을 가져온다.
# Step 04-02: B열에 "1001"값을 텍스트 형식으로 모두 넣는다.
# Step 04-03: C열에 값은 입력하지 않는다.
# Step 04-04: D열에 값은 입력하지 않는다.
# Step 04-05: E열에 값은 "06"값을 텍스트 형식으로 모두 넣는다.
# Step 04-06: F열에 값은 "x"값을 텍스트 형식으로 모두 넣는다.
additional_data_3_1 = {
    'Material': box_16,
    'Plant': ['1001'] * len(box_16),
    'QM proc. active': [''] * len(box_16),
    'QM control key': [''] * len(box_16),
    'Insp type': ['06'] * len(box_16),
    'Active': ['x'] * len(box_16)
}

# Step 05-01: 아래에 다시 A열에 BOX_16 배열 내 값을 가져온다.
# Step 05-02: B열에 "1001"값을 텍스트 형식으로 모두 넣는다.
# Step 05-03: C열에 값은 "x"값을 텍스트 형식으로 모두 넣는다.
# Step 05-04: D열에 값은 "0001"값을 텍스트 형식으로 모두 넣는다.
# Step 05-05: E열에 값은 "0130"값을 텍스트 형식으로 모두 넣는다.
# Step 05-06: F열에 값은 "x"값을 텍스트 형식으로 모두 넣는다.
additional_data_3_2 = {
    'Material': box_16,
    'Plant': ['1001'] * len(box_16),
    'QM proc. active': ['x'] * len(box_16),
    'QM control key': ['0001'] * len(box_16),
    'Insp type': ['0130'] * len(box_16),
    'Active': ['x'] * len(box_16)
}

# Combine all data sets for the third output file
combined_data_3 = {key: data_3[key] + additional_data_3_1[key] + additional_data_3_2[key] for key in data_3}

# Create DataFrame for the third output file
output_df_3 = pd.DataFrame(combined_data_3)

# Save the third output file to Excel
with pd.ExcelWriter(output_file_3, engine='openpyxl') as writer:
    output_df_3.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    
    for col in worksheet.columns:
        for cell in col:
            cell.number_format = "@"
    
    # Apply filter to the table
    tab = Table(displayName="Table3", ref=f"A1:F{len(output_df_3) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    worksheet.add_table(tab)

    # Set column width from A column
    worksheet.column_dimensions['A'].width = 12.3
    worksheet.column_dimensions['B'].width = 9.6
    worksheet.column_dimensions['C'].width = 19.8
    worksheet.column_dimensions['D'].width = 19.5
    worksheet.column_dimensions['E'].width = 13.4
    worksheet.column_dimensions['F'].width = 10.6

# Step 03-01: 네번째 output 파일 이름을 "6-2. Insp_16BOX(수입, 출장검사).xlsx"라고 지정한다.
output_file_4 = f'{save_path}/6-2. Insp_16BOX(수입, 출장검사).xlsx'

# Step 04-01: 네번째 output 파일 A열은 헤더를 "Material"로 입력하고, BOX_16 배열 내 값을 가져온다.
data_4 = {
    'Material': box_16,
    'Plant': ['1001'] * len(box_16),
    'Date': [today_date] * len(box_16),
    'Usage': ['5'] * len(box_16),
    'Status': ['4'] * len(box_16),
    'Dynamic mod level': [''] * len(box_16),
    'Modification rule': [''] * len(box_16),
    'Control key': ['qm02'] * len(box_16),
    'Description': ['수입검사'] * len(box_16),
    'Inspection': ['CHL00001'] * len(box_16),
    'Sampling': ['fo001'] * len(box_16)
}

# Step 05-01: 아래에 다시 A열에 BOX_16 배열 내 값을 가져온다.
# Step 05-02: B열은 헤더를 "Plant"로 입력하고, "1001"값을 텍스트 형식으로 모두 넣는다.
# Step 05-03: C열은 헤더를 "Date"로 입력하고, 값은 오늘날짜를 "YYYYMMDD" 형식으로 모두 넣는다.
# Step 05-04: D열은 헤더를 "Usage"로 입력하고, 값은 "50"값을 텍스트 형식으로 모두 넣는다.
# Step 05-05: E열은 헤더를 "Status"로 입력하고, 값은 "4"값을 텍스트 형식으로 모두 넣는다.
# Step 05-06: F열은 헤더를 "Dynamic mod level"로 입력하고, 값은 입력하지 않는다.
# Step 05-07: G열은 헤더를 "Modification rule"로 입력하고, 값은 입력하지 않는다.
# Step 05-08: H열은 헤더를 "Control key"로 입력하고, 값은 "qm02"값을 텍스트 형식으로 모두 넣는다.
# Step 05-09: I열은 헤더를 "Description"로 입력하고, 값은 "출장검사"값을 텍스트 형식으로 모두 넣는다.
# Step 05-10: J열은 헤더를 "Inspection"로 입력하고, 값은 "CHL00001"값을 텍스트 형식으로 모두 넣는다.
# Step 05-11: K열은 헤더를 "Sampling"로 입력하고, 값은 "fo001"값을 텍스트 형식으로 모두 넣는다.
additional_data_4 = {
    'Material': box_16,
    'Plant': ['1001'] * len(box_16),
    'Date': [today_date] * len(box_16),
    'Usage': ['50'] * len(box_16),
    'Status': ['4'] * len(box_16),
    'Dynamic mod level': [''] * len(box_16),
    'Modification rule': [''] * len(box_16),
    'Control key': ['qm02'] * len(box_16),
    'Description': ['출장검사'] * len(box_16),
    'Inspection': ['CHL00001'] * len(box_16),
    'Sampling': ['fo001'] * len(box_16)
}

# Combine all data sets for the fourth output file
combined_data_4 = {key: data_4[key] + additional_data_4[key] for key in data_4}

# Create DataFrame for the fourth output file
output_df_4 = pd.DataFrame(combined_data_4)

# Save the fourth output file to Excel
with pd.ExcelWriter(output_file_4, engine='openpyxl') as writer:
    output_df_4.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    
    for col in worksheet.columns:
        for cell in col:
            cell.number_format = "@"
    
    # Apply filter to the table
    tab = Table(displayName="Table4", ref=f"A1:K{len(output_df_4) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    worksheet.add_table(tab)

    # Set column width from A column
    worksheet.column_dimensions['A'].width = 12.3
    worksheet.column_dimensions['B'].width = 9.6
    worksheet.column_dimensions['C'].width = 9.3
    worksheet.column_dimensions['D'].width = 10.5
    worksheet.column_dimensions['E'].width = 10.6
    worksheet.column_dimensions['F'].width = 22.9
    worksheet.column_dimensions['G'].width = 21
    worksheet.column_dimensions['H'].width = 15.5
    worksheet.column_dimensions['I'].width = 15.3
    worksheet.column_dimensions['J'].width = 14.3
    worksheet.column_dimensions['K'].width = 13.3

# Step 03-01: 네번째 output 파일 이름을 "6-3. Insp_16BOX(수리검사) (복사 1번).xlsx"라고 지정한다.
output_file_5 = f'{save_path}/6-3. Insp_16BOX(수리검사) (복사 1번).xlsx'

# Step 04-01: 네번째 output 파일 A열은 헤더를 "Material"로 입력하고, BOX_16 배열 내 값을 가져온다.
data_5 = {
    'Material': box_16,
    'Plant': ['1001'] * len(box_16),
    'Date': [today_date] * len(box_16),
    'Usage': ['54'] * len(box_16),
    'Status': ['4'] * len(box_16),
    'Dynamic mod level': [''] * len(box_16),
    'Modification rule': [''] * len(box_16),
    'Control key': ['qm02'] * len(box_16),
    'Description': ['수리검사'] * len(box_16),
    'Inspection': ['CHL00003'] * len(box_16),
    'Sampling': ['to001'] * len(box_16)
}

# Create DataFrame for the fifth output file
output_df_5 = pd.DataFrame(data_5)

# Save the fifth output file to Excel
with pd.ExcelWriter(output_file_5, engine='openpyxl') as writer:
    output_df_5.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    
    for col in worksheet.columns:
        for cell in col:
            cell.number_format = "@"
    
    # Apply filter to the table
    tab = Table(displayName="Table5", ref=f"A1:K{len(output_df_5) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    worksheet.add_table(tab)

    # Set column width from A column
    worksheet.column_dimensions['A'].width = 12.3
    worksheet.column_dimensions['B'].width = 9.6
    worksheet.column_dimensions['C'].width = 9.3
    worksheet.column_dimensions['D'].width = 10.5
    worksheet.column_dimensions['E'].width = 10.6
    worksheet.column_dimensions['F'].width = 22.9
    worksheet.column_dimensions['G'].width = 21
    worksheet.column_dimensions['H'].width = 15.5
    worksheet.column_dimensions['I'].width = 15.3
    worksheet.column_dimensions['J'].width = 14.3
    worksheet.column_dimensions['K'].width = 13.3

# Notify success
print(f"Filtered data saved successfully: \n- {output_file_1}\n- {output_file_2}\n- {output_file_3}\n- {output_file_4}\n- {output_file_5}")
