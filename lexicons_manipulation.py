import pandas as pd

# Load the Excel file
file_path = r'C:\Users\User\Desktop\pp\SDP.xlsx'  # Change this to the path of your file
xls = pd.ExcelFile(file_path)

# Load both sheets into separate DataFrames
sheet1 = pd.read_excel(xls, xls.sheet_names[0])
sheet2 = pd.read_excel(xls, xls.sheet_names[1])

# Convert both DataFrames to sets of words (assuming words are in the first column)
set1 = set(sheet1.iloc[:, 0])
set2 = set(sheet2.iloc[:, 0])

# AND operation (Intersection of both sheets)
and_operation = set1.intersection(set2)
and_df = pd.DataFrame(and_operation, columns=['Common Words'])

# OR operation (Union of both sheets)
or_operation = set1.union(set2)
or_df = pd.DataFrame(or_operation, columns=['All Words'])

# Now save the results into the same Excel file as new sheets
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:  # 'a' mode to append
    and_df.to_excel(writer, sheet_name='AND Operation', index=False)
    or_df.to_excel(writer, sheet_name='OR Operation', index=False)

print("Operations completed and sheets added to the Excel file.")