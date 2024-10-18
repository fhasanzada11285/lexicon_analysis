import pandas as pd

# Load the Excel file
file_path = r'C:\Users\User\Desktop\lexicon_analysis\SDP.xlsx'
excel_data = pd.ExcelFile(file_path)

# Read the sheets into DataFrames
df_lexicon_1 = pd.read_excel(excel_data, sheet_name='lexicon_1')
df_lexicon_2 = pd.read_excel(excel_data, sheet_name='lexicon_2')
df_lexicon_3 = pd.read_excel(excel_data, sheet_name='lexicon_3')

# Standardize emotional columns to numeric values
emotional_columns = ['anger', 'anticipation', 'disgust', 'fear', 'joy', 'negative', 'positive', 'sadness', 'surprise', 'trust']
for df in [df_lexicon_1, df_lexicon_2, df_lexicon_3]:
    for col in emotional_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).apply(lambda x: 1 if x >= 1 else 0)

# AND operation: Words common in all lexicons
common_words = set(df_lexicon_1['English Word']).intersection(df_lexicon_2['English Word'], df_lexicon_3['English Word'])
df_common = pd.concat([
    df_lexicon_1[df_lexicon_1['English Word'].isin(common_words)],
    df_lexicon_2[df_lexicon_2['English Word'].isin(common_words)],
    df_lexicon_3[df_lexicon_3['English Word'].isin(common_words)]
])
df_and_operation = df_common.groupby(['English Word', 'Azerbaijani Word'], as_index=False).max()

# OR operation: All words from all lexicons
df_all = pd.concat([df_lexicon_1, df_lexicon_2, df_lexicon_3])
df_or_operation = df_all.groupby(['English Word', 'Azerbaijani Word'], as_index=False).max()
df_or_operation['common'] = df_or_operation['English Word'].isin(common_words).astype(int)
df_or_operation.sort_values(by=['common', 'English Word'], ascending=[False, True], inplace=True)
df_or_operation.drop(columns='common', inplace=True)

# Save the processed data back to the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    for sheet_name in ['AND Operation', 'OR Operation']:
        if sheet_name in writer.book.sheetnames:
            del writer.book[sheet_name]
    df_and_operation.to_excel(writer, sheet_name='AND Operation', index=False)
    df_or_operation.to_excel(writer, sheet_name='OR Operation', index=False)

print("Excel file has been updated with 'AND Operation' and 'OR Operation' sheets.")
