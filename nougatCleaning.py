#This is the first file of the Nougat OCR to TPC extraction pipeline, cleaning the Nougat markdown files
#by removing blank lines, "repeater" lines, and collapsing tabular files to a single line.
#
#Code/comments written by Mitchell Tague.  Questions?  Email me at tague.mitchell@epa.gov

from suffix_tree import Tree
import os
import glob
import pandas as pd

#Read the list of markdown files, and set up the directory for post-cleaned files.
os.chdir("C:/Users/mtague/OneDrive - Environmental Protection Agency (EPA)/Profile/Documents")
dirfrom = "OCR Repo\\n_results"
dirto = "OCR Repo\\Nougat PostPython"
files = os.listdir(dirfrom)
#Iterate over the files, and read each file's lines.
for file in files:
    f = open(f"{os.getcwd()}\\{dirfrom}\\{file}", "r", encoding='utf-8', errors='ignore')
    lines = f.readlines()
    midlines = []
    #Iterate over the lines.  If a line is just the new line character, we can remove it.
    for line in lines:
        removal = False
        if (line == "\n"):
            removal = True
            next
        #The main issue that certain lines get caught into a repeat loop.  This leads to massive lines,
        #made of a single sentence or table portion repeated ad infinitum.  In order to remove these,
        #we use a suffix tree to identify the substrings with the most repeats.
        tree = Tree()
        tree.add(1, line)
        done = False
        try:
            mrs = tree.maximal_repeats()
        #If the suffix tree doesn't work to even be made, it's a broken line, and we can remove it.
        except:
            removal = True
            next
        #Look to see if any of the substrings are at least 20 characters long, and repeat at least 10
        #times.  The format of the suffix trees are weird, list of characters instead of strings,
        #so to do the count we quickly recreate it as a string.
        else:
            for id, path in mrs:
                if done:
                    break
                if len(path) > 20:
                    tempstr = ""
                    for j in range(0,len(path)):
                        tempstr = tempstr + path[j]
                    if line.count(tempstr) > 10:
                        removal = True
                        done = True
        #If a line isn't marked for removal, add it to the "middle" list of lines to keep.
        if not removal:
            midlines.append(line)
    
    #Now for table cleaning.  "Fixit" is the boolean to say "we are currently in a not-complete" table.
    nextlines = []
    fixit = False
    #Iterate over the lines.
    for line in midlines:
        #If we're in the middle of a table, and the next line has the end of a line, paste it on and append it.
        if fixit and f"\\end{{tabular}}" in line:
            tempstr = f"{tempstr} {line}"
            nextlines.append(tempstr)
            fixit = False
        #If we're in the middle of a table, but there's no new line command at least, the table is broken.
        elif fixit and "&" not in line and "\\\\" not in line:
            fixit = False
        #Otherwise, if we're in the middle of a table, we need to strip the new line character off the next line.
        elif fixit:
            tempstr = tempstr + line.strip('\n')
        #Otherwise, we're not in a table.  We only need to check if we have the start of a table, without the end.
        elif f"\\begin{{tabular}}" in line and f"\\end{{tabular}}" not in line:
            tempstr = line.strip("\n")
            fixit = True
        #Otherwise, the line isn't a table.
        else:
            nextlines.append(line)

    #Last thing we need to handle - sometimes OCR just repeats the same line a lot of times as lines moving down.
    #This is somehow related to the above thing, but structurally different.  Removing these will speed things up.
    curstring = ""
    counter = 1
    start = 0
    dellist = "None"
    for i in range(len(nextlines)):
        #For each line, check if it's equal to the current string.
        line = nextlines[i]
        if line == curstring:
            counter = counter + 1
        #If not, and a line's been repeated at least 5 times, add it to the delete list.
        else:
            if counter > 5:
                if dellist == "None":
                    dellist = list(range(start,i))
                else:
                    for j in range(start,i):
                        dellist.append(j)
            #If it's a new line, it's a new line.  Reset counter and starting index.
            curstring = line
            counter = 1
            start = i
    #Remove the deleted lines.
    if dellist != "None":
        endlines = [j for i,j in enumerate(nextlines) if i not in dellist]
    else:
        endlines = nextlines

    #File's all cleaned up, now to write it with a PP attached.
    with open(f"{os.getcwd()}\\{dirto}\\PP-{file}", "w", encoding = 'utf-8', errors="replace") as output:
        for line in endlines:
            output.write(line)