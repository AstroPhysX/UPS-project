import camelot
import matplotlib.pyplot as plt
import pandas as pd
import re
import datetime
import PyPDF2

def find_df_with_string(dataframes,string):
    result = []
    
    for df in dataframes:
        if isinstance(df, pd.DataFrame) and not df.empty:
            first_row = df.iloc[0]
            if any(first_row.astype(str).str.contains(string)):
                result.append(df)
    
    return result

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