import pandas as pd
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
import glob
import datetime
import os
from pathlib import Path
from tkinter import messagebox
import math
from xml.sax.handler import ContentHandler
from xml.sax import make_parser
import xml.dom.minidom

# version
tool_version = "Arca CSVtoXML 5.0.0"

# pyinstaller --onefile --noconsole --icon convert.ico CSVtoXMLv5.py
#update 2020-06-07 
#prettified using xml.dom.minidom
#revised method that deletes blank columns
#updated for use with MMW 18-6 (combined personal and family names)

#update 2020-05-22
#fixed bug (dateIssued duplication)

#update 2020-05-22
#appended _MODS to filenames constructed from the Arca PID

#update 2020-05-21
#removed prettification
#refactored code

#update 2020-05-20
#integrated content models
#fixed bug with prettification
#added cleaning of additional non-utf-8 characters

#update 2020-05-14
#prettifies XML so that label will be updated on ingest.
#still for use with Master Metadata Workbook 18-5 only 
#will not accept metadata with whole-name (combined given name and family name) fields
#gives pop-up error message if CSV is empty or contains incompatible metadata

#update 2020-02-27
#checked for NoneType in contributor given name and subbed empty string if None
#inserted obligatory question marks around "uri" in <identifier type=uri>

#update 2020-02-14
#adjustment to accommodate multiple languages + backward compatibility for CorporateCreator 
#(as well as CorporateCreator_1 and CorporateCreator_2)
#fixed error in getting Corporate Subject from the input CSV
#added URI
# supported content models
    #book
    #large image
    #audio
    #video
    #newspaper (as a whole) 
    #news issue

#pyinstaller command:



#infer desktop
desktopPath = os.path.expanduser("~/Desktop/")
csvname = ""
filelist=['',None]
#----------------------------------------------------------------------
   
def validate(savePath):
    def parsefile(file):
        parser = make_parser()
        parser.setContentHandler(ContentHandler())
        parser.parse(file)
    
    for filename in Path(savePath).rglob('*.xml'):
        try:
            parsefile(str(filename))
    
        except:
            probFiles += "\n" + str(os.path.basename(filename)) + " was not well formed."
    #return probFiles
       
def browse_button1():  
    # Allow user to select a directory and store it in global var
    # called folder_path1

    lbl1['text'] = ""
    csvname =  filedialog.askopenfilename(initialdir = desktopPath,title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))
    filelist[0] = csvname
    lbl1['text'] = csvname
 
def dropNullCols(df):
    nullcols = []
    for col in df.columns:
        notNull = df[col].notna().sum()
        if notNull < 2:
            nullcols.append(col)
    return nullcols 
   
def convert():
    #list of problem files
    probFiles = ''
    download = False
    filename = filelist[0]
    outputFldr = getOutputFolder()

    # Set default output folder
    if (len(outputFldr) < 1) or outputFldr is None:
        outputFldr = "CSVtoXML_Output"
    savePath = os.path.join(desktopPath,outputFldr)
    if not os.path.exists(savePath):    #if folder does not exist
        os.makedirs(savePath)
  
    try:
        df = pd.read_csv(filename,dtype = "string", encoding = 'utf_8')
    except UnicodeDecodeError:
        df = pd.read_csv(filename,dtype = "string", encoding = 'utf_7')
    
    nullcols = dropNullCols(df)
    df.drop(nullcols, axis=1, inplace=True)
 
    df.columns = map(str.lower, df.columns) #accommodate case variations in column headings

    if "BCRDHSimpleObjectPID".lower() in df.columns: #CSV data has been downloaded and resulting XML will correct existing Arca MODS
        download = True
        
    def clean(xmlStr):
        #replace non-utf-8 characters
        newstr = xmlStr.replace("&","&amp;") #ampersand
        newstr = re.sub(u"\u2013", "-", newstr) #en dash
        newstr = re.sub(u"\u2018", "'", newstr) #curly left quotation mark
        newstr = re.sub(u"\u2019", "'", newstr) #curly right quotation mark
        newstr = re.sub(u"\ufffd","", newstr)# weird extraneous char occurring in hedley photographs stripped CSV
        #newstr = re.sub(u"\x9d", '"',newstr)
        return(newstr)

    def getOutputFilename(df, row,download):
        if 'BCRDHSimpleObjectPID'.lower() in df.columns:
            pid = df.loc[row, 'BCRDHSimpleObjectPID'.lower()]
            fileName = pid.replace(":","_") + "_MODS.xml"
        elif 'PID'.lower() in df.columns:
            pid = df.loc[row, 'PID'.lower()]
            fileName = pid.replace(":","_") + "_MODS.xml"     
        else:
            fileName = df.at[row, 'Filename'.lower()]
            if fileName ==None:
                exit()
            fileName = fileName.strip()
            ext = fileName[-4:]
            fileName = fileName.replace(ext,'.xml')
        fileName = os.path.join(savePath,fileName)
        return fileName
 
    for item in df.itertuples():
        #****MODS TITLE INFO****
        xmlString = '<?xml version="1.0" encoding="UTF-8"?><mods xmlns="http://www.loc.gov/mods/v3" xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink">'

        ti = None
        if 'title' in df.columns:
            ti = item.title
            if pd.isna(ti):
                ti = item.issuetitle
        else:
            ti = item.issuetitle
  
        if pd.notna(ti):
            xmlString += '<titleInfo><title>' + ti + '</title></titleInfo>'
        if 'alternativetitle' in df.columns:
            altTi = item.alternativetitle
        
            if pd.notna(altTi):
                xmlString +='<titleInfo type = "alternative"><title>'+altTi + '</title></titleInfo>'
        #****MODS ORGIN INFO****

        xmlString += '<originInfo>'
   
        date = item.datecreated
        
        if date is None:
            date = "n.d."
        xmlString += '<dateIssued keyDate="yes" encoding="w3cdtf">' + date + '</dateIssued>'
        if "publisher_original" in df.columns:
            pub = item.publisher_original
            if pd.notna(pub): #no publisher
                    xmlString += '<publisher>' + pub + '</publisher>'
            if 'publisher_location' in df.columns:
                publoc = item.publisher_location    
                if publoc is not None:
                    xmlString += '<place><placeTerm type="text">' + publoc + '</placeTerm></place>'
         
        xmlString += '</originInfo>'

        #****PERSONAL_SUBJECTS****
        hdgs = df.filter(like="personalsubject").columns
        if len(hdgs) > 0:
            for hdg in hdgs:
                ps = df.at[item.Index,hdg]
                if pd.isna(ps):
                    break 
                else:                    
                    xmlString += '<subject><name type="personal"><namePart>' + ps + '</namePart></name></subject>'
    
        #****CREATORS****
        hdgs = df.filter(regex='^creator[1-9]').columns
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "creator" + str(j)
                pc =  df.at[item.Index,hdg]
                if pd.isna(pc):
                    break
                else:
                    xmlString += '<name type = "personal"><namePart>' + pc + '</namePart><role><roleTerm type="text" authority="marcrelator">creator</roleTerm></role></name>' 
           
        #****CONTRIBUTORS****
        hdgs = df.filter(regex='^contributor[1-9]').columns #find up to 9 personal contributors
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "contributor" + str(j)
                pcon =  df.at[item.Index,hdg]
                if pd.isna(pcon):
                    break
                else:
                    xmlString += '<name type = "personal"><namePart>' + pcon + '</namePart><role><roleTerm type="text" authority="marcrelator">contributor</roleTerm></role></name>' 
         
        #****CORPORATE CONTRIBUTOR****
        hdgs = df.filter(like="corporatecontributor").columns
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "corporatecontributor" + str(j)     
                corpco = df.at[item.Index,hdg]
                if pd.isna(corpco):
                    break
                else:
                    xmlString += '<name type="corporate"><namePart>' + corpco + '</namePart><role><roleTerm type="text" authority="marcrelator">creator</roleTerm></role></name>'
        
            
        #****CORPORATE CREATOR****
        hdgs = df.filter(like="corporatecreator").columns
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "corporatecreator" + str(j)
                corpcr =  df.at[item.Index,hdg]
                if pd.isna(corpcr):
                    break
                else:
                    xmlString += '<name type="corporate"><namePart>' + corpcr + '</namePart><role><roleTerm type="text" authority="marcrelator">creator</roleTerm></role></name>'
      
        #****CORPORATE SUBJECT****
        hdgs = df.filter(like="corporatesubject").columns
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "corporatesubject" + str(j)
                corpsub =  df.at[item.Index,hdg]
                if pd.isna(corpsub):
                    break
                else:
                    xmlString += '<subject><name type="corporate"><namePart>' + corpsub + '</namePart></name></subject>'

        #****PHYSICAL DESCRIPTION (EXTENT & NOTES)****
        if 'extent' in df.columns:
            extent = item.extent
            if pd.notna(extent):
                xmlString += '<physicalDescription><extent>' + extent + '</extent></physicalDescription>'

        #****ABSTRACT/DESCRIPTION****
        if 'description' in df.columns:
            descr = item.description
            if pd.notna(descr):
                xmlString += '<abstract>' + descr + '</abstract>'

        #****TOPICAL SUBJECTS
        hdgs = df.filter(like="topicalsubject").columns
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "topicalsubject" + str(j)
                top = df.at[item.Index,hdg]
                if pd.isna(top):
                    break
                else:
                    xmlString += '<subject><topic>' + top + '</topic></subject>'
     
        #****COORDINATES****
        if 'coordinates' in df.columns:
            coords = item.coordinates
            if pd.notna(coords):
                xmlString += '<subject><geographic><cartographics>' + coords + '</cartographics></geographic></subject>'

        #****GEOGRAPHIC SUBJECT****
        hdgs = df.filter(like="geographicsubject").columns
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "geographicsubject" + str(j)
                geoSub = df.at[item.Index,hdg]
                if pd.isna(geoSub):
                    break
                else:
                    xmlString += '<subject><geographic>' + geoSub + '</geographic></subject>'    
                        
        #****TEMPORAL SUBJECT****
        if 'daterange' in df.columns:
            dtrng = item.daterange
            if pd.notna(dtrng):
                xmlString += '<subject><temporal>' + dtrng + '</temporal></subject>'

        #****LANGUAGE****
        hdgs = df.filter(like="language").columns
        numHdgs = len(hdgs)
        if numHdgs > 0:
            for j in range(1,numHdgs + 1):
                hdg = "language" + str(j)
                lang = df.at[item.Index,hdg]
                if pd.isna(lang):
                    break
                else:
                    xmlString += '<language><languageTerm>' + lang + '</languageTerm></language>' 

        #****NOTES****
        if 'notes' in df.columns:
            notes = item.notes
            if pd.notna(notes):
                xmlString += '<note>' + notes + '</note>'
        
        #****GENRE / GENRE AUTHORITY****
        if 'genre' in df.columns:
            genre = item.genre
            if pd.notna(genre):
                if 'genreauthority' in df.columns:
                    auth = item.genreauthority
                    if pd.notna(auth):
                        xmlString += '<genre authority="' + auth.lower() + '">' + genre + '</genre>'
                    else:
                        xmlString += '<genre>' + genre + '</genre>' 
                else:
                    xmlString += '<genre>' + genre + '</genre>' 
        #****TYPE****
        if 'type' in df.columns:
            type_ = item.type
            if pd.notna(type_):
                xmlString += '<typeOfResource>' + type_ + '</typeOfResource>'
  
        #****INTERNET MEDIA TYPE****
        if 'internetmediatype' in df.columns:
            imt = item.internetmediatype
            if pd.notna(imt):
                xmlString += '<physicalDescription><internetMediaType>' + imt + '</internetMediaType></physicalDescription>'

        #****IDENTIFIERS****
        if 'accessidentifier' in df.columns:
            aI = item.accessidentifier
            if pd.notna(aI):
                xmlString += '<identifier type="access">' + aI + '</identifier>'
                
        if 'localidentifier' in df.columns:        
            lI = item.localidentifier
            if pd.notna(lI):
                xmlString += '<identifier type="local">' + lI + '</identifier>'
        
        if 'URI'.lower() in df.columns:
            uri = item.uri
            if pd.notna(uri):
                xmlString += '<identifier type="uri">' + uri + '</identifier>'
   
        #****CLASSIFICATION****
        if 'classification' in df.columns:
            classif = item.classification
            if pd.notna(classif):
                xmlString += '<classification authority="lcc">' + classif + '</classification>'

        #****SOURCE****
        if 'source' in df.columns:
            source = item.source
            if pd.notna(source):
                xmlString += '<location><physicalLocation>' + source + '</physicalLocation></location>'
        
        #****ISBN****
        if 'isbn' in df.columns:
            isbn = item.isbn
            if pd.notna(isbn):
                xmlString += '<identifier type="isbn">' + isbn + '</identifier>'
          
        #****RIGHTS****
        if 'rights' in df.columns:
            rts = item.rights
            if pd.notna(rts):
                xmlString += '<accessCondition type="use and reproduction" displayLabel="Restricted">' + rts + '</accessCondition>'

        #****RIGHTSSTATEMENT****
        if 'rightsstatement_uri' in df.columns:
            rStmt = item.rightsstatement_uri
            if pd.isna(rStmt):
                rStmt = 'http://rightsstatements.org/vocab/CNE/1.0/'
            xmlString += '<accessCondition type="use and reproduction" displayLabel="Rights Statement">' + rStmt + '</accessCondition>'

        if 'creativecommons_uri' in df.columns:
            cc_uri = item.creativecommons_uri
            if pd.notna(cc_uri):
                xmlString += '<accessCondition type="use and reproduction" displayLabel="Creative Commons license">' + cc_uri + '</accessCondition>'
       
        #****RECORD CREATION DATE / RECORD ORIGIN*****
        xmlString += '<recordInfo><recordOrigin>'+ tool_version +'</recordOrigin><recordCreationDate keyDate="yes" encoding="w3cdtf">' + datetime.datetime.now().strftime('%Y-%m-%d') + '</recordCreationDate></recordInfo>'
        
        #****RELATEDITEM****
        if 'relatedItem_Title'.lower() in df.columns:
            related_item_title = item.relateditem_title
            if pd.notna(related_item_title):
                xmlString += \
                    '<relatedItem type="host"><titleInfo><title>'+ related_item_title +'</title></titleInfo></relatedItem>'
        if 'relatedItem_PID'.lower() in df.columns:
            related_item_pid = item.relateditem_pid
            if pd.notna(related_item_pid):
                xmlString += \
                    '<relatedItem type="host"><identifier type="PID">'+ related_item_pid +'</identifier></relatedItem>'
        
        #****TAIL****
        xmlString += '</mods>'
 
        xmlString = clean(xmlString)
        fileName = getOutputFilename(df, item.Index, download)
        dest = os.path.join(savePath,fileName)
      
        xmlstr = xml.dom.minidom.parseString(xmlString)  # or xml.dom.minidom.parseString(xml_string)
        xml_pretty_str = xmlstr.toprettyxml()
      
        try:
            with open(dest, "w") as f:
                f.write(xml_pretty_str)
        except Exception as e:
            probFiles += "\n" + str(os.path.basename(dest)) + "could not be written: " + str(e)

         
    if len(probFiles) > 0:
        msg = "Unfortunately, the following files either were not well formed\nor could not be written by the program:"
        msg += "--------------------------------------------------------\n"
        msg += " " + probFiles
        messagebox.showinfo(title = 'Conversion Report: Errors', message = msg)
    finalMsg = "Records have been written to the " + outputFldr + " folder on your desktop."
    messagebox.showinfo(title = 'Conversion Finished', message = finalMsg)



def getOutputFolder():
        return(output.get())
probFiles = ''

root = Tk()

root.eval('tk::PlaceWindow %s center' % root.winfo_toplevel())
root.title = "CSV -> XML"
root.configure(background='#DAE6F0')
folder_path = StringVar()

intro = ttk.Label(master=root,text="Arca_CSVtoXML Version 5.0.0", background='#DAE6F0',font="Arial 15 bold")
intro.grid(row=0,column=1,padx=(10,0),sticky='w')
info = ttk.Label(master=root,text="**Use with metadata format 18-6 (single-field personal names)!**\nThis app can be used with audio, book, large image, newspaper,\nnewspaper issue, PDF, and video content models. It can\nproduce XML both for initial ingest and for revision of existing\nArca MODS, and will auto-detect the intended use.", background='#DAE6F0',font="Arial 9 bold")
info.grid(row=1,column=1,padx=(10,0),sticky='w')


srclbl = ttk.Label(master=root,text="\nChoose CSV to convert",background='#DAE6F0',font="Arial 10 italic")
srclbl.grid(row=2,column=1,padx=(10,0),sticky='w')

button1 = ttk.Button(text="Browse", command=browse_button1)
button1.grid(row=3, column=1,  padx=(10,0),sticky='w')

lbl1 = ttk.Label(master=root,background='#DAE6F0',font="Arial 10")
lbl1.grid(row=4, column=1,padx=(10,5), sticky='w')

destlbl = ttk.Label(master=root,text="Enter Name for Output Folder on Desktop",background='#DAE6F0',font="Arial 10 italic")
destlbl.grid(row=6,column=1,padx=(10,5), sticky='w')

output = ttk.Entry(master=root, width = 30)
output.grid(row = 7, column = 1, padx = (10,0), pady=(0,5), sticky = 'sw')

compButton = ttk.Button(text="Generate XML",command=convert)
compButton.grid(row=8, column=1,padx=(10,0),pady=(15,15),sticky='w')

mainloop()