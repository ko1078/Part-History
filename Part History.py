from operator import index
import numpy as np3
import warnings
warnings.simplefilter("ignore", UserWarning)
import pandas as pd
import dataclasses
import fitz
from PyPDF2 import PdfFileWriter, PdfFileReader
from pandas.tseries.offsets import DateOffset
import pyodbc 
import pypyodbc
import win32com.client as win32
import os
from os.path import join
import os.path
import concurrent.futures
from multiprocessing import freeze_support
from pathlib import Path
import time
import shutil
from datetime import date

import glob

Part = []     #blank list of work Orders
WO = str(input('Part Number - '))
Part.append(WO) # these blocks collect the certs from the user based on number of certs told
Part_name = Part[0]
PO_Requested = pd.DataFrame(Part,columns= ['Part'])

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=STAUBCAD\SIGMANEST;'
                      'Database=SNDBase22;'
                      'Trusted_Connection=yes;')

SERVER_NAME = 'STAUBCAD\SIGMANEST'
DATABASE_NAME = 'SNDBase22'

sql_query = "SELECT WoNumber, SheetName, PartName, ProgramName, RevisionNumber, Material, Thickness, Data3, ProgrammedBy, QtyProgram ,ArcDateTime  FROM [dbo].[PartArchive] "

part = pd.read_sql(sql_query, conn) # request SQL  Table from STAUBCAD

sql_query1 = "SELECT SheetName, ProgramName, PrimeCode, HeatNumber,TaskName  FROM [dbo].[StockArchive] "

stock = pd.read_sql(sql_query1, conn) # request SQL  Table from STAUBCAD



part_shortened = part[part['PartName'].isin(PO_Requested['Part'])]          # Removes all un requested Work Orders from the parts list.


stock_shortened = stock[stock['SheetName'].isin(part_shortened['SheetName'])]          # removes all Sheets from the stock list that aernt required for the WO Numbers Requested.



merged_inner = pd.merge(left=stock_shortened, right=part_shortened,how='left', left_on='ProgramName', right_on='ProgramName') # merges the two data frames of the database and the PO Recietps spreadsheet to matching PO_MTL fields.


d = {'QtyProgram':'sum', 'ArcDateTime':'first',}
df_new = merged_inner.groupby(['WoNumber','PrimeCode','Material','ProgramName','Thickness','Data3',"TaskName",'ProgrammedBy','RevisionNumber','HeatNumber'], as_index=False).aggregate(d).reindex(columns=merged_inner.columns)
df_sorted= df_new.sort_values(by='WoNumber', ascending=False)
df_sorted.to_excel(r'C:\Users\GGehring\Documents\Part_'+Part_name+'_History.xlsx', columns=['WoNumber','ProgramName','RevisionNumber','QtyProgram','ArcDateTime','TaskName','Data3','PrimeCode','HeatNumber','Material','Thickness','ProgrammedBy'],index = False)
