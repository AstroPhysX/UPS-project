# %%
import camelot
import matplotlib.pyplot as plt
import pandas as pd
import re
import datetime
import PyPDF2
#%%
filename="A:/Github Repos/UPS-project/SamplePDFs/2304 Trips.pdf"
#%%
pdf=open(filename,'rb')
pagenum=PyPDF2.PdfFileReader(pdf).getNumPages()
pgrange='1-'+str(pagenum)
pdf.close()
pgrange='141'
#%%
tables_stream = camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61','390,573,774,61'],columns=['47.5,92.9,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2','443.4,488.9,531.4,565,594.9,613.1,629.1,644.5,663.2,686.2,713.4,749.1'],strip_text='--\n',split_text=True)
#tables_lattice = camelot.read_pdf(File,pages='3',process_background=True,strip_text='--\n')

#tables_lattice= camelot.read_pdf(filename,pages=pgrange,strip_text='--\n')
#tables_stream=camelot.read_pdf(filename,flavor='stream',pages=pgrange,table_areas=['42.4,466,80.5,314.5','42.4,314,80.5,159.0'],strip_text='--\n')
s=1 
l=0
print(tables_stream[s].df)
#print(tables_stream[n].df)
#camelot.plot(tables_lattice[l],kind='grid')
camelot.plot(tables_stream[s],kind='contour')
#camelot.plot(tables_lattice[l],kind='contour')
#camelot.plot(tables_lattice[l],kind='joint')
#camelot.plot(tables_lattice[l],kind='line')
#camelot.plot(tables_lattice[l],kind='text')
camelot.plot(tables_stream[s],kind='textedge')
plt.show(block=True)
"""
dflist2=[]
for i in range(0,len(tables_stream)):
    SDF=None
    CT=None
    for index, row in tables_stream[i].df.iterrows():
        row_str=' '.join(row.astype(str))
        SDF=re.search(r"SDF  (\d+)",row_str)
        CT=re.search(r"CT: (\d{2}):(\d{2})",row_str)
        
        if SDF:
            dflist2.append(pd.DataFrame(columns=["Line #","CT"]))
            dflist2[i].loc[0,"Line #"]=int(SDF.group(1))
            
        elif CT:
            h=int(CT.group(1))
            min=int(CT.group(2))
            
            if index>5:
                CT2=datetime.timedelta(hours=h, minutes=min)
                dflist2[i].loc[0,"CT"]= CT1+CT2
                
            CT1=datetime.timedelta(hours=h, minutes=min)
            
            
            
        
            
# initialize an empty list to store the resulting dataframes

#%%
trips = []

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
                trips.append((tables[n].df.loc[tripid_index:index-1]).reset_index(drop=True))
                tables[n].df = tables[n].df.loc[index:]
            
            tripid_index=index
            curr_trip_id = row[0]
            prev_trip_id = curr_trip_id
    if tripid_index==0:
        del tables[n].df
    else:
        # append the last part of the dataframe to the result list    
        trips.append(tables[n].df.reset_index(drop=True))
    
print("  Finished Parsing trips")
#%%

trips_list= []
#this is the way I am going to have to do it
n=0
df=pd.DataFrame(columns=["Trip Id", "DH", "Start time", "End time","Layovers","TAFB","Dest"])
trips_list.append(df)

#saving the trip number
trips_list[n].loc[0,"Trip Id"]=int(re.findall('\d+',trips[n].iloc[0,0])[0])

#saving TAFB
hours, minutes=map(int, trips[n].iloc[9,12].split("h"))
trips_list[n].loc[0,"TAFB"]=datetime.time(hour=hours, minute=minutes)

#Saving DH
DH=0

if re.findall( r'\bDH\b',tables[0].df.iloc[5,1]) or re.findall( r'\bCM\b',tables[0].df.iloc[5,1]):
    DH=DH+1

for i in range(len(tables[0].df)-1,-1,-1):
    string=tables[0].df.iloc[i,1]
    if string1!='':
        if re.findall( r'\bDH\b',string) or re.findall( r'\bCM\b',string):
            DH=DH+1
            break
print(DH)
trips_list[n].loc[0,"DH"]=DH

#Destinations






#we can either set specific values in a dataframe = to something and iterate through each thing
trips_list[n].loc[3,"Dest"]="dfw"
#or create an dataframe that contains all the  destinations and concat it 
dest=pd.DataFrame(columns=["Dest"])
dest["Dest"]=['dfw','sdf','dfw','sdf']
trips_list[n]=pd.concat([trips_list[n],dest],axis=1)    
    
        
        




# print the resulting dataframes
#%%

for i, tables_stream[n].df in enumerate(trips):
    print(f"Dataframe {i+1}:")
    print(tables_stream[n].df)
"""