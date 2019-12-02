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

# version
tool_version = "4.3"
# supported content models
    #book
    #large image
    #audio
    #news issue
#infer desktop
desktopPath = os.path.expanduser("~/Desktop/")
csvname = ""
filelist=['',None]
#----------------------------------------------------------------------
   
def validate(savePath):
    probFiles = ""
    def parsefile(file):
        parser = make_parser()
        parser.setContentHandler(ContentHandler())
        parser.parse(file)
    
    for filename in Path(savePath).rglob('*.xml'):
    #for filename in glob.iglob(savePath, recursive=True):
        try:
            parsefile(str(filename))
    
        except:
            probFiles += "\n" + str(os.path.basename(filename))
    return probFiles

       
def browse_button1():  
    # Allow user to select a directory and store it in global var
    # called folder_path1
    lbl1['text'] = ""
    csvname =  filedialog.askopenfilename(initialdir = desktopPath,title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))
    filelist[0] = csvname
    lbl1['text'] = csvname
   
def convert():
    contentMod = "default"
    download = False
    filename = filelist[0]
    outputFldr = getOutputFolder()
    # Set default output folder
    if (len(outputFldr) < 1) or outputFldr == None:
        outputFldr = "CSVtoXML_Output"
    savePath = os.path.join(desktopPath,outputFldr)
    if not os.path.exists(savePath):    #if folder does not exist
        os.makedirs(savePath)
    df = pd.read_csv(filename, dtype = str)
    df = df.where((pd.notnull(df)), None) #convert Pandas NaNs to None
    df.columns = map(str.lower, df.columns)
    numrows = len(df.index)
    if "PID".lower() in df.columns:
        download = True
    if "IssueTitle".lower() in df.columns:
        contentMod = "newsIssue" 
    elif "dateissued_start" in df.columns:
        contentMod = "newsSerial"

    def clean(xmlStr):
        newstr = xmlStr.replace("&","&amp;")
        return(newstr)

    def getOutputFilename(row,download):
        if download:
            pid = df.at[row, 'PID'.lower()]
            if pid is not None:
                fileName = df.at[row, 'PID'.lower()].replace(":","_") + ".xml"
            else:
                print(df.at[row,'filename'])
                fileName = df.at[row, 'filename'][:-3] + "xml"
                
                
                #re.sub(r'\.\w{3,}', '.xml', df.at[row, 'Filename'.lower()])
        else:
            fileName = df.at[row, 'Filename'.lower()]
            if fileName ==None:
                exit()
            fileName = fileName.strip()
            ext = fileName[-4:]
            fileName = fileName.replace(ext,'.xml')
        fileName = os.path.join(savePath,fileName)
        return fileName
    
    ####____News Issue____####
    
    if contentMod == "newsIssue": #if we are converting newspaper issue metadata
        for row in range(0,numrows):
            if df.at[row,'issuetitle'] == None:
                exit()
            xmlString = '<?xml version="1.0" encoding="UTF-8"?><mods xmlns="http://www.loc.gov/mods/v3" xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink">'
            #Issue Title
            issTitle = df.at[row,'issuetitle']
            xmlString += '<titleInfo><title>' + issTitle + '</title></titleInfo>'
            # CHANGED
            date = df.at[row,'datecreated']
            if date != None:
            # CHANGED
                xmlString += '<originInfo><dateIssued encoding="w3cdtf" keyDate="yes">' + date + '</dateIssued></originInfo>' 
            
            #Volume and Issue Info
            vol = str(df.at[row,'volume'])
            iss = str(df.at[row,'issue'])
            if vol != None or iss != None:
                xmlString += '<part>'
                if vol != None:
                    xmlString += '<detail type="volume"><number>' + vol + '</number></detail>'
                if iss != None:
                    xmlString += '<detail type="issue"><number>' + iss + '</number></detail>'
                xmlString += '</part>'
            
            accessID = str(df.at[row,'identifier'])
            if accessID != None:
                xmlString += '<identifier type="access">' + accessID + '</identifier>'
            
            #****RECORD CREATION DATE / RECORD ORIGIN*****
            xmlString += '<recordInfo><recordOrigin>'+ tool_version +'</recordOrigin><recordCreationDate>' + datetime.datetime.now().strftime('%Y-%m-%d') + '</recordCreationDate></recordInfo>'
            xmlString += '</mods>'
            xmlString = clean(xmlString)
            fileName = df.at[row,'filename'][:-3] + "xml"
            dest = os.path.join(savePath,fileName)
            
            try:
                with open(dest, "wb") as f:
                    f.write(xmlString.encode('utf8'))
            except Exception as e:
                print(str(e))
                print("Write of " + fileName + " failed.")
        
        probFiles = validate(savePath)
        if len(probFiles) > 0:
            msg = "Unfortunately, the following files were not well formed:\n"
            msg += "--------------------------------------------------------\n"
            msg += " " + probFiles
            messagebox.showinfo(title = 'XML Check', message = msg)
        finalMsg = "Records have been written to the " + outputFldr + " folder on your desktop."
        messagebox.showinfo(title = 'Conversion Finished', message = finalMsg)
            
    elif contentMod == "newsSerial":
        for row in range(0,numrows):
            if df.at[row,'title'] == None:
                exit()
            xmlString = '<?xml version="1.0" encoding="UTF-8"?><mods xmlns="http://www.loc.gov/mods/v3" xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink">'
            ti = df.at[row, 'Title'.lower()]
            if ti != None:
                xmlString += '<titleInfo><title>' + ti + '</title></titleInfo>'
                
            #****ORIGIN INFO*****    
            xmlString += '<originInfo>'  #open OriginInfo
            # CHANGED
            pub = df.at[row,'publisher_original']
            if pub is not None:
                xmlString += '<publisher>' + pub + '</publisher>'
            loc = df.at[row,'publisher_location']
            if loc is not None:
                xmlString += '<place><placeTerm type="text">' + loc + '</placeTerm></place>'
            freq = df.at[row,"frequency"]
            if freq != None:
                xmlString += '<frequency>' + freq + '</frequency>'
            dateStart = df.at[row,'dateissued_start']
            dateEnd = df.at[row, 'dateissued_end']
            if dateStart != None:
                xmlString += '<dateIssued point="start">'+ dateStart +'</dateIssued>'
                xmlString += '<dateIssued point="end">' + dateEnd + '</dateIssued>'
            xmlString += '</originInfo>' #close OriginInfo

            # CHANGED
            abs = df.at[row,'abstract']
            if abs is not None:
                xmlString += '<abstract>' + abs + '</abstract>'
            # CHANGED
            dtrng = df.at[row,'daterange']
            if dtrng is not None:
                xmlString += '<subject><temporal>' + dtrng + '</temporal></subject>'
                   
            # CHANGED
            intMedia = df.at[row,'internetmediatype']
            if intMedia is not None:
                xmlString += '<physicalDescription><internetMediaType>' + intMedia + '</internetMediaType></physicalDescription>'

            xmlString += '<physicalDescription><form type="basicform" authority="marcform">print</form></physicalDescription>'
            xmlString += '<typeOfResource>text</typeOfResource>'
            lang = df.at[row,'language']
            if lang:
                xmlString += '<language>' + lang + '</language>'
            xmlString += '<genre authority="aat">newspapers</genre>'
            
                            #****RIGHTS****
            rts = df.at[row,'rights']
            if rts != None:
                xmlString += '<accessCondition type="use and reproduction" displayLabel="Restricted">' + rts + '</accessCondition>'

            #****RIGHTSSTATEMENT****
            rStmt = df.at[row,'rightsstatement']
            if rStmt == None:
                rStmt = 'http://rightsstatements.org/vocab/CNE/1.0/'
            xmlString += '<accessCondition type="use and reproduction" displayLabel="Rights Statement">' + rStmt + '</accessCondition>'
            
            
            #****CREATIVECOMMONSURI****
            if 'CreativeCommons_URI'.lower() in df.columns:
                cc_uri = df.at[row, 'CreativeCommons_URI'.lower()]
                if cc_uri != None:
                    xmlString += '<accessCondition type="use and reproduction" displayLabel="Creative Commons license">' + cc_uri + '</accessCondition>'
            
            #****RECORD CREATION DATE / RECORD ORIGIN*****
            xmlString += '<recordInfo><recordOrigin>'+ tool_version +'</recordOrigin><recordCreationDate>' + datetime.datetime.now().strftime('%Y-%m-%d') + '</recordCreationDate></recordInfo>'
            xmlString += '</mods>'
            
            xmlString = clean(xmlString)
            if download:
                fileName = df.at[row,'pid'].replace(":","_") + ".xml"
            else:
                fileName = ti + ".xml"
            dest = os.path.join(savePath,fileName)
            
            try:
                with open(dest, "wb") as f:
                    f.write(xmlString.encode('utf8'))
            except Exception as e:
                print(str(e))
                print("Write of " + fileName + " failed.")
                
        probFiles = validate(savePath)    
        if len(probFiles) > 0:
            msg = "Unfortunately, the following files were not well formed:\n"
            msg += "--------------------------------------------------------\n"
            msg += " " + probFiles
            messagebox.showinfo(title = 'XML Check', message = msg)
        finalMsg = "Records have been written to the " + outputFldr + " folder on your desktop."
        messagebox.showinfo(title = 'Conversion Finished', message = finalMsg)
            
    else:
        psHeadings = df.filter(regex='subject[0-9]+_[gf]+').columns #find SubjectX_Given or ...Family
        numPSubs = math.ceil((len(psHeadings))/2) # of personal subjects
        numRows = len(df.index)

        for row in range(0,numRows):
            if df.at[row,'title'] == None:
                exit()
            
            #****MODS TITLE INFO****
            xmlString = '<?xml version="1.0" encoding="UTF-8"?><mods xmlns="http://www.loc.gov/mods/v3" xmlns:mods="http://www.loc.gov/mods/v3" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink">'
            ti = df.at[row, 'Title'.lower()]
            if ti != None:
                xmlString += '<titleInfo><title>' + ti + '</title></titleInfo>'
            altTi = None
            if 'AlternativeTitle'.lower() in df.columns:
                altTi = df.at[row, 'AlternativeTitle'.lower()]
            if altTi != None:
                xmlString += '<titleInfo type = "alternative"><title>' + altTi + '</title></titleInfo>'
            
            #****MODS ORGIN INFO****
            date = None
            if 'datecreated' in df.columns:
                date = str(df.at[row, 'DateCreated'.lower()])
            pub = None
            pub_location = None
            if date == None:
                date = "n.d."
            if 'Publisher_Original'.lower() in df.columns:
                pub = df.at[row, 'Publisher_Original'.lower()]
            if 'Publisher_Location'.lower() in df.columns:
                pub_location = df.at[row, 'Publisher_Location'.lower()]

            if pub==None: #no publisher
                # CHANGED
                xmlString += '<originInfo><dateIssued encoding="w3cdtf" keyDate="yes">' + date + '</dateIssued></originInfo>'
            else:
                # CHANGED
                #xmlString += '<originInfo eventType="publication"><publisher>' + pub + '</publisher><dateIssued keyDate="yes" encoding="w3cdtf">'
                xmlString += '<originInfo><publisher>' + pub + '</publisher>'
                if pub_location !=None:
                    xmlString += '<place><placeTerm type="text">' + pub_location + '</placeTerm></place>'
                xmlString += '<dateIssued keyDate="yes" encoding="w3cdtf">' + date + '</dateIssued></originInfo>'

            #****PERSONAL_SUBJECTS****
            for i in range (1, numPSubs + 1):
                fnHdg = "Subject" + str(i) + "_Family"
                gnHdg = "Subject" + str(i) + "_Given"
                fnHdg = fnHdg.lower()
                gnHdg = gnHdg.lower()
                fn = df.at[row,fnHdg]
                gn = df.at[row,gnHdg]
                if fn == None and gn == None:
                    break
                else:
                    xmlString += '<subject><name type="personal">'
                    if fn != None:
                        print("Family name is " + fn)
                        xmlString += '<namePart type="family">' + fn + '</namePart>'
                    if gn != None:
                        print("Given name is " + gn)
                        xmlString += '<namePart type="given">' + gn + '</namePart>'

                    xmlString += '</name></subject>'

            # CHANGED
            #****CREATORS****
            for x in range(1, 4):
                family_column = 'Creator' + str(x) + '_Family'
                given_column = 'Creator' + str(x) + '_Given'
                if family_column.lower() in df.columns:
                    creator_family = df.at[row, family_column.lower()]
                    if creator_family is not None:
                        xmlString +=\
                        '<name type = "personal"><namePart type ="family">' + creator_family + '</namePart>'
                        xmlString +=\
                        '<namePart type ="given">' + df.at[row, given_column.lower()] + '</namePart><role><roleTerm type="text" authority="marcrelator">creator</roleTerm></role></name>'

            # CHANGED - I changed this, as cont_family had been put in the given namePart and cont_given in the family namePart. - sh
            #****CONTRIBUTORS****
            for x in range(1, 3):
                cont_family_col = 'Contributor' + str(x) + '_Family'
                cont_given_col = 'Contributor' + str(x) + '_Given'
                if cont_family_col.lower() in df.columns:
                    cont_family = df.at[row, cont_family_col.lower()]
                    if cont_family is not None:
                        xmlString += \
                        '<name type = "personal"><namePart type="family">' + cont_family + '</namePart>'
                        xmlString += \
                        '<namePart type="given">' + df.at[row, cont_given_col.lower()] + '</namePart><role><roleTerm type="text" authority="marcrelator">contributor</roleTerm></role></name>'
            # CHANGED
            #****CORPORATE CONTRIBUTOR****
            if 'CorporateContributor'.lower() in df.columns:
                corpCo = df.at[row, 'CorporateContributor'.lower()]
                if corpCo is not None:
                    xmlString += '<name type="corporate"><namePart>'+corpCo+'</namePart><role><roleTerm type="text" authority="marcrelator">contributor</roleTerm></role></name>'
            #****CORPORATE CREATOR****
            corpCr = None
            if 'CorporateCreator'.lower() in df.columns:
                corpCr = df.at[row, 'CorporateCreator'.lower()]
            if corpCr != None:
                xmlString += '<name type="corporate"><namePart>' + corpCr + '</namePart><role><roleTerm authority="marcrelator">creator</roleTerm></role></name>'

            # CHANGED
            #****CORPORATE SUBJECT****
            for x in range(1, 3):
                corp_subj_col = 'CorporateSubject_' + str(x)
                if corp_subj_col.lower() in df.columns:
                    corp_subj = df.at[row, corp_subj_col.lower()]
                    if corp_subj is not None:
                        xmlString += \
                       '<subject><name type="corporate"><namePart>' + corp_subj + '</namePart><role></name></subject>'


            #****PHYSICAL DESCRIPTION (EXTENT & NOTES)****
            extent = df.at[row, 'Extent'.lower()]
            if extent != None:
                xmlString += '<physicalDescription><extent>' + extent + '</extent></physicalDescription>'

            #****ABSTRACT/DESCRIPTION****
            descr = df.at[row, 'Description'.lower()]
            if descr != None:
                xmlString += '<abstract>' + descr + '</abstract>'

            #****TOPICAL SUBJECTS
            #topicHdgs = df.filter(regex='Subject[0-9]+_Topic').columns
            topicHdgs = df.filter(like="topic").columns
            print(topicHdgs)
            numTopSubs = len(topicHdgs)
            print(str(numTopSubs))
            for j in range(1,numTopSubs + 1):
                topHdg = "Subject" + str(j) + "_Topic"
                topHdg = topHdg.lower() # changed - SH
                top = df.at[row,topHdg]
                if top == None:
                    break
                else:
                    xmlString += '<subject><topic>' + top + '</topic></subject>'

            #****CORPORATE SUBJECT****
            corpSub = None
            corpSub1 = None
            corpSub2 = None
            if 'CorporateSubject'.lower() in df.columns:
                corpSub = df.at[row, 'CorporateSubject'.lower()]
            elif 'CorporateSubject_1'.lower() in df.columns:
                corpSub1 = df.at[row, 'CorporateSubject_1'.lower()]
                if 'CorporateSubject_2'.lower() in df.columns:
                    corpSub2 = df.at[row, 'CorporateSubject_2'.lower()]
            if corpSub != None:
                xmlString += '<subject><name type="corporate"><namePart>' + corpSub + '</namePart></name></subject>'
            if corpSub1 != None:
                xmlString += '<subject><name type="corporate"><namePart>' + corpSub1 + '</namePart></name></subject>'
            if corpSub2 != None:
                xmlString += '<subject><name type="corporate"><namePart>' + corpSub2 + '</namePart></name></subject>'

            #****COORDINATES****
            coords = None
            if 'Coordinates'.lower() in df.columns:
                coords = df.at[row, 'Coordinates'.lower()]
            if coords != None:
                # CHANGED
                xmlString += '<subject><geographic><cartographics>' + coords + '</cartographics></geographic></subject>'

            # CHANGED
            #****GEOGRAPHIC SUBJECT****
            if 'Subject_Geographic'.lower() in df.columns:
                geoSub = df.at[row, 'Subject_Geographic'.lower()]
                if geoSub != None:
                    xmlString += '<subject><geographic>' + geoSub + '</geographic></subject>'
            else:
                geographic_columns = df.filter(regex='subject[0-9]+_geographic').columns
                num_geographic_columns = len(geographic_columns)
                for j in range(1, num_geographic_columns + 1):
                    geographic_column = 'Subject' + str(j) + '_Geographic'
                    geographic_value = df.at[row, geographic_column.lower()]
                    if geographic_value is not None:
                        xmlString += \
                            '<subject><geographic>' + geographic_value + '</geographic></subject>'                
                            
            #****TEMPORAL SUBJECT****
            dtrng = df.at[row,'daterange']
            if dtrng is not None:
                xmlString += '<subject><temporal>' + dtrng + '</temporal></subject>'

            #****LANGUAGE****
            if 'Language'.lower() in df.columns:
                if df.at[row, 'Language'.lower()] != None:
                    lang = df.at[row, 'Language'.lower()]
                    xmlString += '<languageTerm type="text">' + lang + '</languageTerm>'
                else:
                    xmlString += '<languageTerm type="text">English</languageTerm>'

            # CHANGED
            #****GENRE / GENRE AUTHORITY****
            genre = None
            genre_authority = None
            if 'Genre'.lower() in df.columns:
                genre = df.at[row, 'Genre'.lower()]
            if 'GenreAuthority'.lower() in df.columns:
                genre_authority = df.at[row, 'GenreAuthority'.lower()]

            if (genre is not None) and (genre_authority is not None):
                # CHANGED
                #xmlString += '<genre authority="aat">' + g + '</genre>'
                xmlString += '<genre authority="' + genre_authority.lower() + '">' + genre + '</genre>' #changed to put genre auth in lower case -sh

            #****TYPE****
            if 'Type'.lower() in df.columns:
                type_ = df.at[row, 'Type'.lower()]
                if type_ != None:
                    xmlString += '<typeOfResource>' + type_ + '</typeOfResource>'

            #****INTERNET MEDIA TYPE****
            if 'internetMediaType'.lower() in df.columns:
                imt = df.at[row, 'internetMediaType'.lower()]
                if imt != None:
                    # CHANGED
                    xmlString +=  \
                        '<physicalDescription><internetMediaType>' + imt + '</internetMediaType></physicalDescription>'


            #****IDENTIFIERS****'
            aI = df.at[row, 'AccessIdentifier'.lower()]
            lI = str(df.at[row, 'LocalIdentifier'.lower()])
            if aI !=None:
                xmlString += '<identifier type="access">' + aI + '</identifier>'
            if lI != None:
                xmlString += '<identifier type="local">' + lI + '</identifier>'
            if 'URI'.lower() in df.columns:
                uri = df.at[row, 'URI'.lower()]
                if uri != None:
                    xmlString += '<identifier type="uri">' + uri + '</identifier>'

            #****CLASSIFICATION****
            if 'Classification'.lower() in df.columns:
                classif = df.at[row, "Classification".lower()]
                if classif != None:
                    xmlString += '<classification authority="lcc">' + classif + '</classification>'

            #****SOURCE****
            source = df.at[row, 'Source'.lower()]
            if source != None:
                xmlString += '<location><physicalLocation>' + df.at[row, 'Source'.lower()] + '</physicalLocation></location>'

            #****ISBN****
            # CHANGED
            if 'ISBN'.lower() in df.columns:
                isbn = df.at[row, 'ISBN'.lower()]
                if isbn is not None:
                    xmlString += '<identifier type="isbn">' + isbn + "</identifier>"

            #****RIGHTS****
            rts = df.at[row, 'Rights'.lower()]
            if rts != None:
                xmlString += '<accessCondition type="use and reproduction" displayLabel="Restricted">' + rts + '</accessCondition>'

            #****RIGHTSSTATEMENT****
            rStmt = df.at[row, 'RightsStatement'.lower()]
            if rStmt == None:
                rStmt = 'http://rightsstatements.org/vocab/CNE/1.0/'
            xmlString += '<accessCondition type="use and reproduction" displayLabel="Rights Statement">' + rStmt + '</accessCondition>'
            

            #****CREATIVECOMMONS_URI****
            if 'CreativeCommons_URI'.lower() in df.columns:
                cc_uri = df.at[row, 'CreativeCommons_URI'.lower()]
                if cc_uri != None:
                    xmlString += '<accessCondition type="use and reproduction" displayLabel="Creative Commons license">' + cc_uri + '</accessCondition>'

            # CHANGEd
            #****RECORD CREATION DATE / RECORD ORIGIN*****
            xmlString += '<recordInfo><recordOrigin>'+ tool_version +'</recordOrigin><recordCreationDate>' + datetime.datetime.now().strftime('%Y-%m-%d') + '</recordCreationDate></recordInfo>'
            # CHANGED
            #****RELATEDITEM****
            if 'relatedItem_Title'.lower() in df.columns:
                related_item_title = df.at[row, 'relatedItem_Title'.lower()]
                if related_item_title is not None:
                    xmlString += \
                        '<relatedItem type="host"><titleInfo><title>'+ related_item_title +'</title></titleInfo></relatedItem>'
            if 'relatedItem_PID'.lower() in df.columns:
                related_item_pid = df.at[row, 'relatedItem_PID'.lower()]
                if related_item_pid is not None:
                    xmlString += \
                        '<relatedItem type="host"><identifier type="PID">'+ related_item_pid +'</identifier></relatedItem>'


            #****TAIL****
            xmlString += '</mods>'
            xmlString = clean(xmlString)
            fileName = getOutputFilename(row, download)
            dest = os.path.join(savePath,fileName)
            
            try:
                with open(dest, "wb") as f:
                    f.write(xmlString.encode('utf8'))
            except Exception as e:
                print(str(e))
                print("Write of " + fileName + " failed.")
        
        probFiles = validate(savePath)    
        if len(probFiles) > 0:
            msg = "Unfortunately, the following files were not well formed:\n"
            msg += "--------------------------------------------------------\n"
            msg += " " + probFiles
            messagebox.showinfo(title = 'XML Check', message = msg)
        finalMsg = "Records have been written to the " + outputFldr + " folder on your desktop."
        messagebox.showinfo(title = 'Conversion Finished', message = finalMsg)


def getOutputFolder():
        return(output.get())

root = Tk()

root.eval('tk::PlaceWindow %s center' % root.winfo_toplevel())
root.title = "CSV -> XML"
root.configure(background='#DAE6F0')
folder_path = StringVar()

intro = ttk.Label(master=root,text="Arca_CSVtoXML Version 4.3", background='#DAE6F0',font="Arial 15 bold")
intro.grid(row=0,column=1,padx=(10,0),sticky='w')
info = ttk.Label(master=root,text="This app can be used with audio, book, large image, newspaper,\nnewspaper issue, PDF, and video content models.", background='#DAE6F0',font="Arial 9 bold")
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