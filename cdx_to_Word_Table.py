# cdx_to_Word_Table.py
# By: Marc Garard 2023 Aug 18
# Basic method of inserting ChemDraw Object into a Word Document Table starting with SDF format
# imports
from win32com import client
import pandas as pd
import os

# controls
# absolute paths are required.  Relative paths fail.
dirPath = "C:/workingDir/" # makes this your path where the folders & files are.  Ex: "C:/Users/<username>/cdxToWordTable"
fileDocSavePath = dirPath + "word/tabled.docx"
sdfPath = dirPath + "sdf/"
cdxPath = dirPath + "cdx/"

def sdfToCDX( df, IDcol, sdfDir, cdxDir  ):
    # Open ChemDraw APP Only
    chemdraw = client.Dispatch("ChemDraw.Application") # connect to ChemDraw
    chemdraw.Visible = False
    filelist = []
    for row in df.iloc:
        cd = chemdraw.Documents.Open( sdfDir + row[ IDcol ] + ".sdf" ) # open sdf and convert to cdx
        cdxPath = cdxDir + row[ IDcol ] + ".cdx"
        cd.SaveAs( cdxPath )
        cd.Close()
        filelist.append( cdxPath )
    # Close ChemDraw
    chemdraw.Quit()
    return filelist

def addCDXToTable( table, filepath, row, col = 2 ):
    table.Cell( Row = row, Column = col ).Range.InlineShapes.AddOLEObject( ClassType="ChemDraw.Document.6.0", FileName=filepath  )
    return

def addIDToTable( table, text, row, col = 1 ):
    table.Cell( Row = row, Column = col ).Range.Text = text
    return

# Initalize dataframe with ID data for files
IDs = [ name.split('.')[0] for name in os.listdir( sdfPath ) ]
df = pd.DataFrame( { 'ID' : IDs } )

# Generate cdx files from sdfs
df[ 'cdxfile' ] = df.apply( lambda x: sdfToCDX( df, 'ID', sdfPath, cdxPath ), axis = 0 ) # convert sdf to cdx. NOTE: df still a series, axis = 0

# Open Word
word = client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add() # new document
# Create table
wdRange = doc.Content
wdRange.Tables.Add( doc.Range(), df.shape[0] + 1, 2) # create table with proper dimensions. +1 for header

# Format Table
table = doc.Tables.Item(1)
table.AutoFormat( 36 )
idCol = table.Columns(1) # ID as column 1
molCol = table.Columns(2) # molecule as column 2
idCol.Cells(1).Range.Text = "ID" # Column Headers
molCol.Cells(1).Range.Text = "ChemDraw Mol" # Column Headers

df = df.reset_index( drop = True ) # clean index
df[ 'DataRow' ] = df.index + 2 # sets desired row of insertion

# insert data into table
try: 
    df.apply( lambda x: addIDToTable( table = table, text = str( x.ID ), row = x.DataRow, col = 1 ), axis = 1 ) # insert ID column values
    df.apply( lambda x: addCDXToTable( table = table, filepath = str( x.cdxfile ), row = x.DataRow, col = 2 ), axis = 1 ) # insert cdx object to word
except: pass

# save and close
doc.SaveAs( fileDocSavePath + "newWord.docx" )
doc.Close(False)
word.Quit()