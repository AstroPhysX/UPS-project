import camelot
import matplotlib.pyplot as plt
import pandas as pd
import re
import datetime
import PyPDF2
import threading

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

def find_df_with_string(dataframes,string):
    result = []
    
    for df in dataframes:
        if isinstance(df, pd.DataFrame) and not df.empty:
            first_row = df.iloc[0]
            if any(first_row.astype(str).str.contains(string)):
                result.append(df)
    
    return result

def trip_extract(filename,dflist):  
    print("Extracting Trips")
    pdf=open(filename,'rb')
    lastpage=PyPDF2.PdfFileReader(pdf).getNumPages()
    startpage=1
    pgrange=str(startpage)+'-'+str(lastpage)
    pdf.close()
    
    print("  Obtained the page range: ",pgrange)
    print("  Reading pdf. Getting Trips this may take a few minutes...")
    
    while True:
        try:
            tables=camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61','390,573,774,61'],columns=['47.5,67.8,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2','443.4,464.2,531.4,565,594.9,613.1,629.1,644.5,663.2,686.2,713.4,749.1'],strip_text='--\n',split_text=True)
            break
        except:
            startpage=startpage+1
            pgrange=str(startpage)+'-'+str(lastpage)
    
    print("  Done reading pdf")
    print("  Parsing Trips")
    
    for n in range(0,len(tables)):
    # initialize variables to keep track of the current and previous trip ids
        tripid_index=0
        prev_trip_id = None
        curr_trip_id = None
        # loop through each row of the dataframe
        for index, row in tables[n].df.iterrows():
            # check if the current row contains a trip id
    
            if 'Trip Id:' in row[0]:
                # if it does, split the dataframe into two parts
                if prev_trip_id is not None:
                    dflist.append((tables[n].df.loc[tripid_index:index-1]).reset_index(drop=True))
                    tables[n].df = tables[n].df.loc[index:]
            
                tripid_index=index
                curr_trip_id = row[0]
                prev_trip_id = curr_trip_id
    if tripid_index==0:
       del tables[n].df
    else:
        # append the last part of the dataframe to the result list    
        dflist.append(tables[n].df.reset_index(drop=True))
    
    print("  Finished Parsing trips")

def line_extract(filename,dflist1,dflist2):
    print("Extracting lines")
    
    pdf=open(filename,'rb')
    startpg=2
    lastpg=PyPDF2.PdfFileReader(pdf).getNumPages()
    beforelastpg=lastpg-1
    pgrange=str(startpg)+'-'+str(beforelastpg)
    pdf.close()
    
    print("  Obtained the page range: ",pgrange)
    print("  Reading pdf. Getting Lines this may take a few minutes....")
    
    tables=[]
    while True:
        try:
            extract=camelot.read_pdf(filename,flavor='stream',pages=pgrange,table_areas=['79.6,468,779.7,384.8','42.4,466,80.5,314.5','79.6,313,779.7,230.98','79.6,384.8,779.7,314.0','79.6,230.98,779.7,159.0','42.4,314,80.5,159.0'],columns=['104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','42.4','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','42.4'],strip_text='--\n')
            for i in range(0,len(extract)):
                tables.append(extract[i].df)
            break
        except:
            startpg+=1
            pgrange=str(startpg)+'-'+str(beforelastpg)
            print("  Updated page range:",pgrange)
    try:
        pgrange=str(lastpg)
        extractlast=camelot.read_pdf(filename,flavor='stream',pages=pgrange,table_areas=['79.6,468,779.7,384.8','42.4,466,80.5,314.5','79.6,313,779.7,230.98','79.6,384.8,779.7,314.0','79.6,230.98,779.7,159.0','42.4,314,80.5,159.0'],columns=['104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','42.4','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','42.4'],strip_text='--\n')
        for i in range(0,len(extractlast)):
                tables.append(extractlast[i].df)
    except:
        extractlast=camelot.read_pdf(filename,flavor='stream',pages=pgrange,table_areas=['79.6,468,779.7,384.8','42.4,466,80.5,314.5','79.6,384.8,779.7,314.0'],columns=['104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','42.4','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3'],strip_text='--\n')
        for i in range(0,len(extractlast)):
            tables.append(extractlast[i].df)
    
    print("  Done reading pdfs")        
    #cut here
    print("  Parsing lines")
    #Parsing Lines
    for i in range(0,int(len(tables)),6):
        
        no_com_table=tables[i].loc[1:].reset_index(drop=True)
        
        if no_com_table.iloc[0,0] in ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"] and tables[i+2].iloc[0,0] in ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]:
            df=pd.concat([tables[i].loc[1:].reset_index(drop=True),tables[i+2]],axis=1,ignore_index=True)
            dflist1.append(df)
            try:
                df=pd.concat([tables[i+3].loc[1:].reset_index(drop=True),tables[i+4]],axis=1,ignore_index=True)
                dflist1.append(df)
            except:
                break
            
    for i in range (0,len(dflist1)):   
        
        if len(dflist1[i].index)==5:
            empty_row=pd.DataFrame({}, index=[2.5])
            dflist1[i]=pd.concat([dflist1[i].loc[:2],empty_row, dflist1[i].loc[3:]]).reset_index(drop=True)
            dflist1[i]=dflist1[i].fillna(' ')
            
        dflist1[i].rename(index={0:"Day",1:"Date",2:"trip #",3:"Destination",4:"Time",5:"# Hours"},inplace=True)
    print("  Lines obtained")
    
    #Paring line info
    print("  Parsing lineinfo")
    lineinfo=find_df_with_string(tables,"SDF")
    PP1=find_df_with_string(tables,"Comment:")
    
    for i in range(0,len(lineinfo)):
        emptydf=pd.DataFrame(columns=["Line #","CT","Comment:"])
        dflist2.append(emptydf)
        comment=re.search(r'Comment:(.*)',str(PP1[i].loc[0,0]))
        
        for index, row in lineinfo[i].iterrows():
            row_str=' '.join(row.astype(str))
            SDF=re.search(r"SDF  (\d+)",row_str)
            CT=re.search(r"CT: (\d{2}):(\d{2})",row_str)
            
            if SDF:
                dflist2[i].loc[0,"Line #"]=int(SDF.group(1))
                dflist2[i].loc[0,"Comment:"]=str(comment.group(1))
            elif CT:
                h=float(CT.group(1))
                min=float(CT.group(2))
                CT=h + min/60
                if index>5:
                    CT2= CT
                    dflist2[i].loc[0,"CT"]= CT1+CT2
                CT1=CT
    print("  Lineinfo obtained")
                

linesfile="A:/Github Repos/UPS-project/SamplePDFs/2304 Lines.pdf"
tripsfile="A:/Github Repos/UPS-project/SamplePDFs/2304 Trips.pdf"
trips=[]
lines=[]
lineinfo=[]
trip_extract(tripsfile,trips)
line_extract(linesfile,lines,lineinfo)
