"""
The following is the Extractor script for the 3M Document set.  This script took the Nougat results after cleaning (see the NougatCleaner
script for more information), identifies all tables and parameters (important values within the text), and extracts them into a single
table with the most relevant data extractable - document, chemical, table or parameter, page number, description, and a value and units
for parameters.  In addition, the script separately extracts tables into their own excel files, using Nougat's spatial formatting to 
attempt recreating the original table.

All code written by Mitchell Tague.  Questions?  Email me at tague.mitchell@epa.gov
"""

import pandas as pd
import os
import glob
import re
import tqdm
from thefuzz import fuzz

#pageFinder - a helper function to use old OCR results in order to identify which page a line appears on.
def pageFinder(slug, lines, j, table = False):
    #Point to the old file directory, and use the slug to identify the file in question.
    os.chdir("C://Users//mtague//OneDrive - Environmental Protection Agency (EPA)//Profile//Documents//OCR Repo//OCR Repo//txt_results")
    filenamestart = f"OCRResults_{slug}*"
    oglines = []
    #Use glob to open the correct file from the namestart using the slug.  Open file and read lines.
    for file in glob.glob(filenamestart):
        f = open(f"{os.getcwd()}\\{file}", "r", encoding='utf-8', errors='ignore')
        oglines = f.readlines()
    #Tables were not captured properly in the original text, so instead we're gonna use the lines around it to identify the page.
    if table:
        starttext = lines[max(j-1,0)]
        endtext = lines[min(j+1,len(lines)-1)]
        startpage = 0
        endpage = 0
        startline = j
    else:
        text = lines[j]
    #If the line is too short, then attach the lines around it to make it long enough.
    big = False
    if not table and len(text) < 50:
        " ".join(lines[(j-1):j+1])
        big = True
    #Set up page trackers for running through the old file, as well as a "best match" placeholder
    temppage = 1
    page = 0
    matchscore = 0
    #Iterate over the lines.  Identify places where OCR put page splits, and use them to update the page tracker.
    for i in range(1, len(oglines)):
        templine = oglines[i].strip()
        if templine[0:5] == "PAGE:":
            try: 
                temppage = int(templine[6:])
            except:
                next
            next
        #Find matches.  Use thefuzz to identify a 0-100 matching score between two strings, and if the line is a better match than any
        #previous, update the page and the matchscore.
        if not big and not table:
            tempscore = fuzz.partial_ratio(templine, text)
            if tempscore > matchscore:
                matchscore = tempscore
                page = temppage
        #When big, we similarly should join the lines in the original document.
        if big and not table:
            templine = " ".join(oglines[(i-1):(i+2)])
            tempscore = fuzz.partial_ratio(templine, text)
            if tempscore > matchscore:
                matchscore = tempscore
                page = temppage
        #Table has to do starting line and ending line separately.
        if table:
            tempscore = fuzz.partial_ratio(templine, starttext)
            if tempscore > matchscore:
                matchscore = tempscore
                startpage = temppage
                startline = i
    #Once starting line is done, we know ending line has to be after it.  This will help catch for false positives.
    if table:
        endpage = startpage
        for i in range(startline+1, len(oglines)):
            templine = oglines[i]
            tempscore = fuzz.partial_ratio(templine, endtext)
            if tempscore > matchscore:
                matchscore = tempscore
                endpage = temppage
        #If a table has two pages, we need to set that up as a returnable result.  Otherwise, we're done.
        if startpage == endpage:
            page = startpage
        else:
            page = [startpage, endpage]
    return page

#This function is for creating a row in the final results table for parameters identified by a nearby unit.
def unitRow(words, j):
    #List out the parameters we're looking for, for identification.
    paramslist = ["concentration", "uptake", "clearance", "biodegredation", "absorption", 
              "adsorption", "solubility", "recovery", "coefficient", "lod", "loq",
              "limit of detection", "limit of quantitation", "reaction rate", "rate constant",
              "half life", "halflife", "half-life", "desorption", "cod", "bod",
              "chemical oxygen demand", "biochemical oxygen demand", "dissolved o2",
              "dissolved oxygen", "bioconcentration", "bcf", "photodegredation", "inhibition",
              "ec50","lc50","ld50","ed50","detection limit","quantitation limit", 
              "lethal dose", "lethal concentration", "ic50", "noael", "loael", "tlv", "constant",
              "rfd", "reference dose", "screening", "threshhold", "ael", "log p", "logp", "log-p"
              ]
    #Set up holders for the unit, the number, and the parameter.
    unit = words[j]
    number = "not found"
    param = "not found"
    #Look through the preceding few words to try to identify a number.
    for k in range(max((j-7),0), j):
        word = words[k]
        if word.replace(".", "").isnumeric():
            number = word
            break
    #Do the exact same thing for parameters, using multi-word parsing for multi-word parameters.
    for k in range(min((j+3),len(words)-1), 1, -1):
        word = words[k]
        if word in paramslist:
            param = word
            break
        twowords = " ".join(words[(k-1):(k+1)])
        if twowords in paramslist:
            param = twowords
            break
        threewords = " ".join(words[(k-2):(k+1)])
        if threewords in paramslist:
            param = threewords
            break
    #If no parameter is found in the run through exactly, run it again, using fuzz to do a partial matching scheme similar to pageFinder.
    if param == "not found":
        matchscore = 0
        for term in paramslist:
            longth = len(term.split(" "))
            if longth == 1:
                for k in range(max((j-10),0), min((j+3),len(words))):
                    temp = fuzz.ratio(term, words[k])
                    if temp > matchscore & temp > 80:
                        matchscore = temp
                        param = term
            if longth == 2:
                for k in range(max((j-10),0), min((j+3),len(words))):
                    twowords = " ".join(words[(k-1):(k+1)])
                    temp = fuzz.ratio(term, twowords)
                    if temp > matchscore & temp > 80:
                        matchscore = temp
                        param = term
            if longth == 3:
                for k in range(max((j-10),0), min((j+3),len(words))):
                    threewords = " ".join(words[(k-2):(k+1)])
                    temp = fuzz.ratio(term, threewords)
                    if temp > matchscore & temp > 80:
                        matchscore = temp
                        param = term
    #Return a list set up like a row of the eventual document.
    object = ["", "", 1, "PAGE", param, number, unit]
    return object
    
#An easy function, for those parameters that cannot be identified via unit.  Uses the identification of a parameter, and looks for
#a numeric around it.
def unitlessRow(words, term):
    param = term
    number = "not found"
    for word in words:
        if word.replace(".", "").isnumeric():
            number = word
            break
    return ["", "", 1, "PAGE", param, number, "none"]

#Helper function for the tableMaker.  Tackles rows that have a /multicolumn{} included, by parsing the multicolumns to correctly identify
#what column certain terms should be in.
def multisolver(line, rows):
    #Strip any whitespace from both sides, and set up your result list.
    line = line.strip()
    thisRow = [""]*rows
    #If there's a table beginning here, the overall line is broken, get out of the multisolver, investigate.
    if f"\\begin{{tabular}}" in line:
        print(f"{slug} sucks at {i}")
        return thisRow
    #Set up a column indexer (i), find the index of the line of the next ampersand and "\multicolumn", set up the terminator boolean
    #We specifically use " & " because it helps not identifying ampersands within multicolumns
    i = 0
    multind = line.find("multicolumn")
    ampind = line.find(" & ")
    multend = False
    #While loop
    while not multend:
        #If the next \multicolumn comes before the next ampersand, we need to parse it.
        if (multind < ampind or ampind == -1) and multind != -1:
            #If there's no ampersand in the future, this is the last run of the loop.
            if ampind == -1:
                multend = True
            #Identify the indices of open and close brackets within the line, before the ampersand.
            opens = [i for i, ltr in enumerate(line[:ampind]) if ltr == "{"]
            closes = [i for i, ltr in enumerate(line[:ampind]) if ltr == "}"]
            #Try to use these open and closes to find the number of columns the multicolumn covers.  If you can't, structural issues, fix manually.
            try:
                num = int(re.findall(r'\d+',line[opens[0]:closes[0]])[0])
            except:
                print(f"{slug} sucks at {i}")
                return "None"
            #Identify the last index for opens.  Use the number found above, and find its average, to find what index to put the multicolumn in.
            k=len(opens)-1
            thisind = (num - 1) // 2 + i
            #If we have as many opens as closes, use the last pair of that to identify the string inside multi.  If not, use ampind to at least
            #capture the important stuff with as little excess as possible.
            if len(opens) == len(closes):
                tempstr = line[(opens[k]+1):(closes[k]-1)]
            else:
                tempstr = line[(opens[k]+1):(ampind-1)]
            #Add to the row.  If the index in question is beyond the listed number of rows, use append to tack it on.
            if thisind >= rows:
                thisRow.append(tempstr)
            else:
                thisRow[thisind] = tempstr
            #Identify the new line using ampind, strip of whitespace, find the new multind, ampind, and update the column tracker.
            line = line[(ampind+2):]
            line = line.strip()
            multind = line.find("multicolumn")
            ampind = line.find("&")
            i = i + num
        #If an ampersand is next, it's way easier.
        elif (ampind < multind or multind == -1) and ampind != -1:
            #If there's no multicolumn in the future, yadda yadda.
            if multind == -1:
                multend = True
            #No parsing needed.  The line is the line.  Go ahead and shove it into the row.
            thisRow[i] = line[:ampind]
            #Update the line, multind and ampind, and the column tracker.
            line = line[(ampind+2):]
            line = line.strip()
            multind = line.find("multicolumn")
            ampind = line.find("&")
            i = i + 1
        #If it's not one of these two scenarios, we're done.
        else:
            multend = True
    return thisRow

#Naming helper function for the files we store tables as. Uses recursion to identify the first free path we can make. 
def tableName(filename, pages, dirto, j):
    if j == 0:
        string = f"{filename}_PAGE_{pages}.xlsx"
    else:
        string = f"{filename}_PAGE_{pages}_{j}.xlsx"
    if os.path.isfile(f"{dirto}//{string}"):
        string = tableName(filename, pages, dirto, j+1)
    return string

#This is a miniature version of the tableMaker function, that doesn't export to excel files, just creates a new row in the table/parameter
#maker.  Used for rerunning the TPCs in particular due to some small specific issues.  Fairly self explanatory and built from the bones of
#tableMaker, so I'm not going to comment in depth.    
def minitableMaker(lines, slug, i):
    line = lines[i]
    lastletter = line.find("}",16)
    lastspot = lastletter + 1
    line = line[lastspot:]
    templine = line
    newspot = templine.find("\\\\")
    templine = templine[:newspot]
    newspotl = line.find("\\\\")
    line = line[(newspotl+2):]
    lastspot = newspotl
    desc = templine
    pages = pageFinder(slug, lines, i, table = True)
    return ["", "", 0, pages, desc, "NA", "NA"]

#The function that handles any tables in the file, both in creating the row for the TPC tracking and the excel file of the table.
def tableMaker(lines, file, slug, i):
    #Identify the line in question.  Tables in markdown will always start with \begin{tabular}{...}, with the ... being combinations of
    #"l", "c", and "r".  These are supposed to represent the columns of the table (and their alignments), and thus initially we used them
    #to find the number of columns involved.  We soon found out this led to plenty of errors, so we added a hefty number of bonus columns,
    #and have checks in place to handle things if even more columns are used.  I called this variable "rows" because I am dumb. 
    line = lines[i]
    lastletter = line.find("}",16)
    letters = [j for j in line[16:lastletter] if j in ["l", "c", "r"]]
    rows = len(letters) + 10
    #Booleans for whether we're done with the table, and whether we're on the first row, to identify descriptions.
    end = False
    first = True
    #Use the last } we found earlier to cut off the "\begin{tabular}" start.  Build the result container list for the excel table, and an
    #iteration tracker to prevent infinite hellscapes.
    lastspot = lastletter + 1
    line = line[lastspot:]
    result = []
    iter = 0
    while not end:
        #When caught in an infinite hellscape, break out of it and fix it manually.
        if iter >= 50:
            print(f"{slug} sucks at {i}")
            return "None."
        iter = iter + 1
        #Identify the upcoming line.  Markdown uses "\\" to end a row, which we have to represent as "\\\\" for special character escaping.
        #Cut it off at that line, take the templine, and set up the rest of the table line for the future.
        templine = line
        newspot = templine.find("\\\\")
        templine = templine[:newspot]
        newspotl = line.find("\\\\")
        line = line[(newspotl+2):]
        lastspot = newspotl
        #If the full line is empty space, we're done.
        if re.search("[ \t]+", line):
            end = True
            next
        #If this is the first line, use it to pull the description for the TPC table. 
        if first:
            desc = templine
            first = False
        #Look for the end of a table ("\end{tabular}") in order to cut it off and prime the while loop to close.
        tabular = templine.find("\\end")
        if tabular != -1:
            end = True
            templine = templine[:tabular]
        #Look for "\hline", a little piece that tells markdown tables where to draw horizontal lines, but do nothing to benefit us.
        hhere = True
        while hhere:
            hline = templine.find("\\hline")
            if hline != -1:
                templine = f"{templine[:hline]}{templine[(hline+6):]}"
            else:
                hhere = False
        #If that removes everything, continue on to the next row.
        if templine.isspace():
            continue
        #If multicolumn is part of the row, punt it to multisolver.
        if "multicolumn" in templine:
            result.append(multisolver(templine,rows))
        #Otherwise, we can proceed normally.  Check for ampersands, the markdown tabular code for "next column," and send a blank line if none
        #are present. 
        else: 
            rowcheck = templine.count("&")
            if rowcheck == 0:
                thisRow = [""]*rows
                ind = (rows - 1) // 2
                thisRow[ind] = templine
                result.append(thisRow)
            #Otherwise, we can begin parsing.  Set up another while loop, and look for the next ampersand.
            else:
                amplineend = False
                ampind = templine.find("&")
                j = 0
                thisRow = [""]*rows
                #If you're not at the end of the ampersands, add the line up to the ampersand, cut it off, and repeat.
                while not amplineend:
                    thisRow[j] = templine[:ampind]
                    j = j + 1
                    templine = templine[(ampind+1):]
                    ampind = templine.find("&")
                    if ampind == -1:
                        amplineend = True
                #When we're done, add one last line on, and append the row to the main result.
                if j < rows:
                    thisRow[j] = templine
                else:
                    thisRow = thisRow.append(templine)
                result.append(thisRow)
    #Try to turn the result into a table.  If we can't, result is empty, so there is no table.
    try:
        result = pd.DataFrame(result)
    except:
        return "None"
    #Find the pages, take the file name by cutting off mmd.
    pages = pageFinder(slug, lines, i, table = True)
    filename = file[:-4]
    #If you have two pages from pageFinder, turn it into a string.
    if isinstance(pages, list):
        pages = f"{pages[0]}-{pages[1]}"
    #Use tableName to find the name of the string, push the table to excel, and return a row for the TPC table.
    dirto = "C://Users//mtague//OneDrive - Environmental Protection Agency (EPA)//Profile//Documents//Python Scripts//3M Scripts//Extractor//Tables"
    string = tableName(filename,pages,dirto,0)
    result.to_excel(f"{dirto}//{string}")
    return ["", "", 0, pages, desc, "NA", "NA"]

#OKAY MAIN FILE TIME.  Set up the working directory, the list of units, and the list of parameters to care about without units.
os.chdir("C://Users//mtague//OneDrive - Environmental Protection Agency (EPA)//Profile//Documents//Python Scripts//3M Scripts//Extractor")
 
unitslist = ["ppm", "ppb", "%", "percent", "per cent", "mg", "kg", "µCi", "µg", 
             "g", "L", "m3", "mL", "µM", "amu", "day", "d", "days", "hr", "hour", "lbs",
             "ng", "year", "years", "hours", "mpb", "yr", "ppt", "m2", "h", "uCi", "ug", "degrees", "F",
             "C", "K", "parts per million", "parts per billion"]
unitless = ["log p", "logp", "log-p", "reaction rate", "rate constant"]

#Identify files and get the slugs from the directory, pull our list of chemicals from a separate excel, and set up our TPC table list.
dirfrom = "C://Users//mtague//OneDrive - Environmental Protection Agency (EPA)//Profile//Documents//OCR Repo//Nougat PostPostPython"
files = os.listdir(dirfrom)
fileslugs = [i[3:7] for i in files]
chemslist = pd.read_excel("Chemical List.xlsx").Chemical
for i in range(len(chemslist)):
    chemslist[i] = str(chemslist[i]).lower()
objectlist = []

#Iterate over the files.  Indexing used to run through the slugs as well, and tqdm provides a progress bar.
for m in tqdm.trange(len(files)):
    #Pull the file and the slug.  Open the file and read it.  Set up chemical tracker.
    file = files[m]
    slug = fileslugs[m]
    f = open(f"{dirfrom}\\{file}", "r", encoding='utf-8', errors='ignore')
    lines = f.readlines()
    currentchem = "None Found"
    #Read one line at a time.
    for i in range(len(lines)):
        line = lines[i]
        #If a line has a table in it, push it to table maker.  String results exist in table maker to be skippable here.
        if f"\\begin{{tabular}}" in line:
            object = tableMaker(lines, file, slug, i)
            if isinstance(object,str):
                next
            #Add the filename and the current chemical, and append this to the TPC result listing.
            else:
                object[0] = file
                object[1] = currentchem
                if len(objectlist) == 0:
                    objectlist = [object]
                else:
                    objectlist.append(object)
        #If not a table, split the line into words to search for parameters.  Iterate over the list of words.
        else: 
            words = line.split(" ")
            for j in range(len(words)):
                object = False
                word = words[j].lower()
                #If the word is in the chemical list, update current chemical.  try/except there so that I don't have to debug 20 different
                #word fits that wouldn't be included anyways.
                try:
                    hmuhph = chemslist.eq(word).any()
                    if hmuhph:
                        currentchem = word
                        next
                except:
                    pass
                #If the word is a unit, look at it like it's for a parameter.  If neither a parameter nor a value are found, skip it.
                if word in unitslist:
                    object = unitRow(words, j)
                    if object[4] == "not found" and object[5] == "not found":
                        next
                    #Identify the file name, current chemical, and page, and make a TPC row out of it.
                    object[0] = file
                    object[1] = currentchem
                    object[3] = pageFinder(slug, lines, i)
                    if len(objectlist) == 0:
                        objectlist = [object]
                    else:
                        objectlist.append(object)
                #This works similar to above, but splits a word on /s in order to find combo units (aka g/mL).
                elif "/" in word:
                    test = word.split("/")
                    for testword in test:
                        if testword in unitslist:
                            object = unitRow(words, j)
                            if object[4] == "not found" and object[5] == "not found":
                                next
                            object[0] = file
                            object[1] = currentchem
                            object[3] = pageFinder(slug, lines, i)
                            if len(objectlist) == 0:
                                objectlist = [object]
                            else:
                                objectlist.append(object)
            #Run the line back if no parameter or table is found, to search for unitless parameters afterwards.  Works similarly to above,
            #and runs checks for both one word and two word combinations.
            if object == False:
                for j in range(len(words)):
                    word = words[j]
                    if word in unitless:
                        object = unitlessRow(words, word)
                        if object[5] == "not found":
                            next
                        object[0] = file
                        object[1] = currentchem
                        object[3] = pageFinder(slug, lines, i)
                        if len(objectlist) == 0:
                            objectlist = [object]
                        else:
                            objectlist.append(object)
                else:
                    twowords = " ".join(words[(j-1):(j+1)])
                    if twowords in unitless:
                        object = unitlessRow(words, twowords)
                        if object[5] == "not found":
                            next
                        object[0] = file
                        object[1] = currentchem
                        object[3] = pageFinder(slug, lines, i)
                        if len(objectlist) == 0:
                            objectlist = [object]
                        else:
                            objectlist.append(object)
    #Every 500 files, boot the current list of tables and parameters to a separate excel file and clear it out.  This is a memory reduction
    #procedure to speed up massive document runs.  Push these to excel, and after the loop is done, push the rest to excel.
    if m % 500 == 0:
        objectframe = pd.DataFrame(objectlist)
        os.chdir("C://Users//mtague//OneDrive - Environmental Protection Agency (EPA)//Profile//Documents//Python Scripts//3M Scripts//Extractor")
        objectframe.to_excel(f"{os.getcwd()}\\Big Results Chems {m}.xlsx")
        objectlist = []
objectframe = pd.DataFrame(objectlist)
os.chdir("C://Users//mtague//OneDrive - Environmental Protection Agency (EPA)//Profile//Documents//Python Scripts//3M Scripts//Extractor")
objectframe.to_excel(f"{os.getcwd()}\\Big Results Chems Rest.xlsx")