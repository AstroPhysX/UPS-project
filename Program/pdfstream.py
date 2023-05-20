import camelot
import matplotlib.pyplot as plt
import pandas as pd
import re
import datetime

def check_duplicate_columns(df):
    """
    Checks if a pandas dataframe has duplicate column names and adds a comma to the second occurrence of the column name.
    """
    # Get a list of all column names
    columns = list(df.columns)

    # Create a dictionary to keep track of how many times each column name appears
    column_counts = {}

    # Loop through all column names and count how many times each one appears
    for column in columns:
        if column in column_counts:
            column_counts[column] += 1
        else:
            column_counts[column] = 1

    # Loop through all column names again and add a comma to the second occurrence of each one
    for i, column in enumerate(columns):
        if column_counts[column] > 1:
            if i != columns.index(column):
                columns[i] = column + " "

    # Update the column names in the dataframe
    df.columns = columns

    return df

def export_to_excel(all_lines, filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    all_lines.to_excel(writer, sheet_name="Sheet1", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    (max_row, max_col) = all_lines.shape
    column_settings = []
    for header in all_lines.columns:
        column_settings.append({'header': header})

    worksheet.add_table(0, 0, max_row, max_col-1, {'columns': column_settings})
    writer.close()

def main():
    File= "2304 Lines.pdf"
    n=0

    linepp1 = camelot.read_pdf(File,flavor='stream',pages='3',table_areas=['79.6,456,779.7,384.8','79.6,302.0,779.7,230.98'],columns=['104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3'],strip_text='--\n')
    linepp2 = camelot.read_pdf(File,flavor='stream',pages='3',table_areas=['79.6,384.8,779.7,314.0','79.6,230.98,779.7,159.0'],columns=['104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3'],strip_text='--\n')
    lineinfo= camelot.read_pdf(File,flavor='stream',pages='3',table_areas=['39,469,81,315','41.9,314,80.5,159.0'],strip_text='--\n')


    lines=[]
    for n in range(0,len(linepp1)):


        if (len(linepp1[n].df.index))==5:
            empty_row=pd.DataFrame({}, index=[2.5])

            linepp1[n].df=pd.concat([linepp1[n].df.loc[:2],empty_row, linepp1[n].df.loc[3:]]).reset_index(drop=True)
            linepp1[n].df=linepp1[n].df.fillna(' ')

        if(len(linepp2[n].df.index))==5:
            empty_row=pd.DataFrame({}, index=[2.5])

            linepp2[n].df=pd.concat([linepp2[n].df.loc[:2],empty_row, linepp2[n].df.loc[3:]]).reset_index(drop=True)
            linepp2[n].df=linepp2[n].df.fillna(' ')

        linepp1[n].df=pd.concat([linepp1[n].df,linepp2[n].df],axis=1)
        linepp1[n].df.columns=(linepp1[n].df.iloc[0] + ' ' + linepp1[n].df.iloc[1])
        linepp1[n].df.rename(index={0:"Day",1:"Date",2:"trip #",3:"Destination",4:"Time",5:"# Hours"},inplace=True)
        linepp1[n].df=linepp1[n].df.iloc[2:]
        lines.append(linepp1[n].df)
    
    
    print(lines[1])
        
    
    

    all_lines=linepp1[0].df

    for n in range(1,len(linepp1)):
        all_lines=pd.concat([all_lines,linepp1[n].df])

    all_lines=check_duplicate_columns(all_lines)

    lineinfo[n].df=lineinfo[n].df.applymap(lambda x: re.sub('[a-zA-Z\s]', '', str(x)))
    lineinfo[n].df=lineinfo[n].df.applymap(lambda x: re.sub(':', '', str(x),1))
    lineinfo[n].df.rename(index={0:"Line #",1:"??",2:"PP",3:"CT"})
    print(lineinfo[0].df)

    #export_to_excel(all_lines, "test.xlsx")

    plt.show(block=True)

if __name__=="__main__":
    main()