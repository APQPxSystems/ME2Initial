import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter

# Streamlit configurations
st.set_page_config(page_title="ExcelComparison | KentK.", layout="wide")
hide_st_style = """
                <style>
                #MainMenu {visibility:hidden;}
                footer {visibility:hidden;}
                header {visibility:hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html=True)

# Remove top white space
st.markdown("""
        <style>
            .block-container {
                    padding-top: 0rem;
                    padding-bottom: 5rem;
                    padding-left: 2rem;
                    padding-right: 2rem;
                }
        </style>
        """, unsafe_allow_html=True)

# Sidebar Configurations
with st.sidebar:
    st.write("## How to Use:")
    st.write("#### 1. Drag and drop excel file of old data (must be xlsx).")
    st.write("#### 2. Drag and drop excel file of new data (must be xlsx).")
    st.write("#### 3. Scroll down for the preview of differences and download button.")
    st.write("#### 4. Click the dowload button to download the differences as excel file.")
    st.write("#### 5. In the downloaded excel file, highlighted column titles indicate that there are differences on their values.")
    st.write("__________________________________")
    st.write("### Kent Katigbak | Systems Engineering")

st.title("Excel Comparison App")
st.write("This web app compares the contents of two excel files and generates another excel file containing the unique values of each original file.")
st.write("__________________________")

# File Uploader
old_col, new_col = st.columns([1, 1])

with old_col:
    old_excel = st.file_uploader("Upload old excel file.", type="xlsx")
    if old_excel is not None:
        st.write("Preview of old excel file:")
        old_excel = pd.read_excel(old_excel)
        st.dataframe(old_excel)
    else:
        st.write("Please upload old excel file.")

with new_col:
    new_excel = st.file_uploader("Upload new excel file.", type="xlsx")
    if new_excel is not None:
        st.write("Preview of new excel file:")
        new_excel = pd.read_excel(new_excel)
        st.dataframe(new_excel)
    else:
        st.write("Please upload new excel file.")
st.write("__________________________")

if old_excel is not None and new_excel is not None:
    # Align the differences on columns
    diff = new_excel.compare(old_excel, keep_shape=True, keep_equal=False)
    st.write("Differences between the data of the two files:")
    st.dataframe(diff)
    st.write('''Note 1: The contents of each SELF column are from the new excel file
            while the contents of each OTHER column are from the old excel file.''')
    st.write('''Note 2: Cells with values indicate the values with updated data from old excel file to new excel file.
            NONE means there are no changes.''')

    # Flatten the differences DataFrame
    def flatten_diff(diff_df):
        flattened_df = pd.DataFrame()
        for col in diff_df.columns.levels[0]:
            self_col = diff_df[col]['self']
            other_col = diff_df[col]['other']
            flattened_df[f'{col}_self'] = self_col
            flattened_df[f'{col}_other'] = other_col
        return flattened_df

    flattened_diff = flatten_diff(diff)

    # Function to create an Excel file with differences and highlight column titles
    def create_excel_with_diff(flattened_diff_df):
        output = BytesIO()
        
        # Create an Excel writer object using xlsxwriter
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write the flattened differences DataFrame to the Excel file
            flattened_diff_df.to_excel(writer, sheet_name='Differences', index=True)
            
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Differences']
            
            # Define the format for highlighted headers
            header_format = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C'})
            
            # Iterate through the columns to check for non-null values
            for col_num, column in enumerate(flattened_diff_df.columns, 1):
                # Check if there are any non-null values in the column
                if flattened_diff_df[column].notnull().any():
                    # Apply the header format to the column header
                    worksheet.write(0, col_num, column, header_format)
            
            writer.close()
        output.seek(0)
        return output

    st.write("__________________________")

    # Add Download button
    excel_data = create_excel_with_diff(flattened_diff)
    st.download_button(
        label="Download Differences as Excel",
        data=excel_data,
        file_name='differences.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
