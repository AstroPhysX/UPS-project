import camelot
import matplotlib.pyplot as plt
import pandas as pd
import re
import datetime
import PyPDF2

filename="A:/Jerome/Python Project/SamplePDFs/2304 Lines.pdf"
#filename="A:/Jerome/Python Project/SamplePDFs/2304 Trips.pdf"

pdf=open(filename,'rb')
lastpage=PyPDF2.PdfFileReader(pdf).getNumPages()
startpage=3
pgrange=str(startpage)+'-'+str(lastpage)
pdf.close()

pgrange='55'
tables_lattice= camelot.read_pdf(filename,pages=pgrange,strip_text='--\n')
tables_stream=camelot.read_pdf(filename,flavor='stream',pages=pgrange,table_areas=['79.6,456,779.7,384.8','42.4,466,80.5,314.5','79.6,302,779.7,230.98','79.6,384.8,779.7,314.0','79.6,230.98,779.7,159.0','42.4,314,80.5,159.0'],columns=['104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','42.4','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','104.6,128.8,152.9,177.1,201.2,225,249.5,273.4,297.6,321.6,345.6,369.8,393.9,418.2,442.1,466.4,490.4,514.3,538.6,562.7,586.8,611,635,659.1,683.4,707.3,731.3','42.4'],strip_text='--\n')
s=0  
l=0

print(tables_stream[s].df)
#print(tables_lattice[l].df)

#camelot.plot(tables_lattice[l],kind='grid')
#camelot.plot(tables_stream[s],kind='contour')
#camelot.plot(tables_lattice[l],kind='contour')
#camelot.plot(tables_lattice[l],kind='joint')
#camelot.plot(tables_lattice[l],kind='line')
#camelot.plot(tables_lattice[l],kind='text')
#camelot.plot(tables_stream[s],kind='textedge')
plt.show(block=True)