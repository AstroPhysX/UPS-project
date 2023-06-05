import camelot
import matplotlib.pyplot as plt
import pandas as pd
import re
import PyPDF2

def separate_trips(tables):
    trips = []  # initialize the list to store the separated trips
    
    for df in tables:
        # initialize variables to keep track of the current and previous trip ids
        tripid_index = 0
        prev_trip_id = None
        curr_trip_id = None
        tripid_count=0
        # loop through each row of the dataframe
        for index, row in df.iterrows():
            # check if the current row contains a trip id
            if 'Trip Id:' in row[0]:
                # if it does, split the dataframe into two parts
                if prev_trip_id is not None:
                    trips.append(df.loc[tripid_index:index-1].reset_index(drop=True))
                    df = df.loc[index:]  

                if tripid_count==0 and index!= 0:
                    tripid_index=index
                    
                tripid_index = index
                curr_trip_id = row[0]
                prev_trip_id = curr_trip_id
                tripid_count+=1
        
        if tripid_count == 0:
            del df
        else:
            # append the last part of the dataframe to the result list    
            trips.append(df.loc[tripid_index:].reset_index(drop=True))
    
    return trips

def search_text_in_pdf(file_path, search_string, start_page=0, backward_search=False):
    with open(file_path, 'rb') as file:
        pdf = PyPDF2.PdfFileReader(file)
        num_pages = pdf.numPages

        if backward_search:
            page_range = range(num_pages - 1, start_page - 1, -1)
        else:
            page_range = range(start_page, num_pages)

        for page_number in page_range:
            page = pdf.getPage(page_number)
            text = page.extractText()

            if search_string in text:
                if backward_search:
                    return page_number
                else:
                    return page_number

        return None  # Return None if the search string is not found in the specified range

def process_trips(tripsdf, dflist,open=False):
    for n in range(len(tripsdf)):
        emptydf = pd.DataFrame(columns=["Trip Id", "DH", "Start time", "End time", "Layovers", "TAFB", "Dest"])
        dflist.append(emptydf)
        o=0
        if open:
          o=1
          print(tripsdf[n].iloc[9-o, 12])
        # Adding the Trip Id
        dflist[n].loc[0, "Trip Id"] = int(re.findall('\d+', tripsdf[n].iloc[0, 0])[0])

        # Adding TAFB
        hours, minutes = tripsdf[n].iloc[9-o, 12].split("h")
        dflist[n].loc[0, "TAFB"] = int(hours) + (int(minutes) / 60)

        # Adding number of DH/CM at end and beginning
        DH = 0
        
        if re.findall(r'\bDH\b', tripsdf[n].iloc[5-o, 1]) or re.findall(r'\bCM\b', tripsdf[n].iloc[5-o, 1]):
            DH+= 1
        for i in range(len(tripsdf[n]) - 1, -1, -1):
            string = tripsdf[n].iloc[i, 1]
            notemptyline=re.search(r'[1-9A-Z]', string)
            if notemptyline:
                if re.findall(r'\bDH\b', string) or re.findall(r'\bCM\b', string):
                    DH += 1
                    break
                else:
                    break
            
        dflist[n].loc[0, "DH"] = DH

        # Adding Layovers and destinations
        k = 0
        df = tripsdf[n].loc[4-o:, 2].reset_index(drop=True)
        for i in range(0, len(df) - 1):
            emptyline=re.findall(r'[A-Z]{3}', df.loc[i])
            branch1 = re.findall(r'[A-Z]{3}', df.loc[i+1])
            if branch1 and not emptyline:
                branch1 = re.findall(r'[A-Z]{3}', df.loc[i + 1])
                branch2 = re.findall(r'[A-Z]{3}', df.loc[i + 2])
                branch3 = re.findall(r'[A-Z]{3}', df.loc[i + 3])
                if branch1 and branch2 and branch3:
                    dflist[n].loc[k, "Dest"] = branch1[0] + '-' + branch1[1] + '-' + branch2[1] + '-' + branch3[1]
                    dflist[n].loc[k, "Layovers"] = branch3[1]
                    k += 1
                elif branch1 and branch2:
                    dflist[n].loc[k, "Dest"] = branch1[0] + '-' + branch1[1] + '-' + branch2[1]
                    dflist[n].loc[k, "Layovers"] = branch2[1]
                    k += 1
                else:
                    dflist[n].loc[k, "Dest"] = branch1[0] + "-" + branch1[1]
                    dflist[n].loc[k, "Layovers"] = branch1[1]
                    k += 1

        # Adding the start and end times of trip
        df = tripsdf[n].loc[4-o:, 3:4].reset_index(drop=True)
        
        # Finding start time
        for i in range(len(df)):
            start = re.findall(r'\d{2}:\d{2}', df.loc[i, 3])
            if start:
                dflist[n].loc[0, "Start time"] = start[0]
                break

        # Finding end time
        for i in range(len(df) - 1, 0, -1):
            end = re.findall(r'\d{2}:\d{2}', df.loc[i, 4])
            if end:
                dflist[n].loc[0, "End time"] = end[0]
                break

def trip_extract(filename,tables,open_tables):
    #Fidning trips
    print("Extracting Trips")
    pdf=open(filename,'rb')
    startpage=2
    lastpage=search_text_in_pdf(filename,"Open Trips Report",backward_search=True)
    beforelast=lastpage-1
    pgrange=str(startpage)+'-'+str(beforelast)
    pdf.close()   
    print("  Obtained the page range: ",pgrange)
    print("  Reading pdf. Getting line Trips this may take a few minutes...")   
    
    while True:
        try:
            extract=camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61','390,573,774,61'],columns=['47.5,92.9,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2','443.4,488.9,531.4,565,594.9,613.1,629.1,644.5,663.2,686.2,713.4,749.1'],strip_text='--\n',split_text=True)
            for i in range(0,len(extract)):
                tables.append(extract[i].df)
        
            break
        except:
            startpage=startpage+1
            pgrange=str(startpage)+'-'+str(lastpage)
            print("  Updated page range:",pgrange)
    try:
        pgrange=str(lastpage)
        extractlast=camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61','390,573,774,61'],columns=['47.5,92.9,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2','443.4,488.9,531.4,565,594.9,613.1,629.1,644.5,663.2,686.2,713.4,749.1'],strip_text='--\n',split_text=True)
        for i in range(0,len(extractlast)):
            tables.append(extractlast[i].df)
    except:
        extractlast=camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61'],columns=['47.5,92.9,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2'],strip_text='--\n',split_text=True)
        for i in range(0,len(extractlast)):
            tables.append(extractlast[i].df)
    
    print("  Done reading line trips")
    
    #Finding open trips
    print("Extracting Open Trips")
    pdf=open(filename,'rb')
    startpage=lastpage+1
    lastpage=search_text_in_pdf(filename,"Trips to Flight Report",start_page=startpage)
    beforelast=lastpage-1
    pgrange=str(startpage)+'-'+str(beforelast)
    pdf.close()
    
    while True:
        try:
            extract=camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61','390,573,774,61'],columns=['47.5,92.9,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2','443.4,488.9,531.4,565,594.9,613.1,629.1,644.5,663.2,686.2,713.4,750'],strip_text='--\n',split_text=True)
            for i in range(0,len(extract)):
                open_tables.append(extract[i].df)
            break
        except:
            startpage=startpage+1
            pgrange=str(startpage)+'-'+str(lastpage)
            print("  Updated page range:",pgrange)
    try:
        pgrange=str(lastpage)
        extractlast=camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61','390,573,774,61'],columns=['47.5,92.9,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2','443.4,488.9,531.4,565,594.9,613.1,629.1,644.5,663.2,686.2,713.4,750'],strip_text='--\n',split_text=True)
        for i in range(0,len(extractlast)):
            open_tables.append(extractlast[i].df)
    except:
        extractlast=camelot.read_pdf(filename,pages=pgrange,flavor='stream',table_areas=['13.4,573,390,61'],columns=['47.5,92.9,136.1,168.9,196.6,214.7,231.2,249.2,265,288.7,317.5,353.2'],strip_text='--\n',split_text=True)
        for i in range(0,len(extractlast)):
            open_tables.append(extractlast[i].df)
    
    print("  Done readng open lines")
    print("  Done reading pdf")
    
def get_trips(filename, trips, opentrips):
    tables=[]
    open_tables=[]
    
    trip_extract(filename,tables,open_tables)
    raw_trips=separate_trips(tables)
    raw_open=separate_trips(open_tables)
    process_trips(raw_trips,trips)
    process_trips(raw_open,opentrips,open=True)
    
    
    #open trips need readjusting for zones