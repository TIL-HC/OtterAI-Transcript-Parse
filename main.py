import streamlit as st
import re
import pandas as pd
from io import StringIO, BytesIO
import xlsxwriter

#Set page configuration
st.set_page_config(
    page_title="OtterAI Transcript to Excel",
    page_icon="📄",
)

#Set header
st.title("🎤➡📄OtterAI Transcript to Excel")
st.markdown('This app takes **OtterAI** transcript **.txt** files and converts them to Excel sheets. When exporting transcripts from OtterAI, ensure that both **"Combine"** \
            options are **unticked**.  Upload OtterAI .txt transcripts using the "Drag and drop files here" section below. You can select and upload **multiple files** at once. \
            All transcripts will be output as **one Excel workbook** in **separate sheets**. The **name of the .txt file** will be the **name of the sheet** in the \
            Excel workbook. Sheet names can only be **31 characters** in length (longer will be cut off). Duplicate sheet names will replace the final 4 characters of \
            the name with " (duplicate number)".')

#Get file
uploaded_files = st.file_uploader("Select files", type='txt', accept_multiple_files=True, label_visibility='hidden')


#Create function that reads a txt file and creates dataframe
def create_trns_df(file):
    
    #Read file line by line
    stringio = StringIO(file.getvalue().decode('utf-8'))
    string_dta = stringio.readlines()

    #Set variables for loop and saving of text
    cnt = 0
    ls1 = []
    ls2 = []

    #Loop through each line, save every 1st and 2nd row in separate lists
    for line in string_dta:
        if cnt == 0:
            ls1.append(line)
            cnt += 1
        elif cnt == 1:
            ls2.append(line)
            cnt += 1
        else:
            cnt = 0

    #Reset counter and create list to store full data
    cnt = 0
    dta = []

    #Loop through lists of text taken from lines and create lists with complete data
    for item in ls1:
        row = []
        rgx = re.search("(.+)\s\s(\d.+)", item)
        if rgx:
            row.append(cnt+1)
            row.append(rgx.group(1).strip())
            row.append(rgx.group(2).strip())
            row.append(ls2[cnt].replace('\n','').strip())
            dta.append(row)
        cnt += 1

    #Put list of data into pandas dataframe
    df = pd.DataFrame(dta, columns = ['ConverationID', 'Name', 'Time', 'Text'])

    #Return dataframe
    return df

#Create Excel workbook
output = BytesIO()
workbook = xlsxwriter.Workbook(output, {'in_memory': True})

#Check if files uploaded
if uploaded_files:
    #Start loading spinner
    with st.spinner('Processing files...'):
        
        #Loop through every file and create data frame and sheet name
        sheet_counts = {}
        for file in uploaded_files:
            file.seek(0)
            df = create_trns_df(file)
            sheet_name = file.name.replace('.txt', '')[:31]  # Remove the file extension from the sheet name

            # Check if the sheet name already exists in the dictionary
            if sheet_name in sheet_counts:
                sheet_counts[sheet_name] += 1
                if len(sheet_name) > 25:
                    sheet_name = f"{sheet_name[:26]} ({sheet_counts[sheet_name]})"
                else:
                    sheet_name = f"{sheet_name} ({sheet_counts[sheet_name]})"
            else:
                sheet_counts[sheet_name] = 0

            #Add sheet to workbook with sheet name
            worksheet = workbook.add_worksheet(sheet_name)

            #Add headers to first row of sheet
            worksheet.write('A1', 'ConversationID')
            worksheet.write('B1', 'Name')
            worksheet.write('C1', 'Time')
            worksheet.write('D1', 'Text')

            # Write the DataFrame data to Excel starting from row 2
            for row_num, (index, row) in enumerate(df.iterrows(), start=1):
                worksheet.write(row_num, 0, row['ConverationID'])
                worksheet.write(row_num, 1, row['Name'])
                worksheet.write(row_num, 2, row['Time'])
                worksheet.write(row_num, 3, row['Text'])

        #Close the workbook once all sheets are added
        workbook.close()
    
    #Show download button to get Excel workbook
    st.download_button(
        label="Download Excel workbook",
        data=output.getvalue(),
        file_name="transcripts_parsed.xlsx",
        mime="application/vnd.ms-excel"
    )
        