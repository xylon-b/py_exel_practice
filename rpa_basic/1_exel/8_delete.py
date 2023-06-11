from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

#ws.delete_rows(8) # 8번째 줄에있는 7번학생 데이터 삭제
#ws.delete_rows(8, 3) #8번째 줄부터총 3줄 삭제

#wb.save("sample_delete_rows.xlsx")

#ws.delete.cols(2) # 2번쨰 열 (B)
ws.delete_cols(2, 2) #2번쨰 영로부터 총 2개 열 삭제 

wb.save("sample_delete_col.xlsx")