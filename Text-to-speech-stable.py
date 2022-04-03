#This is a program to make an audiobook version of every file in the directory

import pyttsx3, sys, os, pathlib, PyPDF2, time

sys.path.append('C://Users//Username//Desktop//E-books')

engine = pyttsx3.init() # object creation

""" RATE"""
rate = engine.getProperty('rate')   # getting details of current speaking rate
engine.setProperty('rate', 200)     # setting up new voice rate


"""VOLUME"""
volume = engine.getProperty('volume')   #getting to know current volume level (min=0 and max=1)
engine.setProperty('volume',1.0)    # setting up volume level  between 0 and 1

"""VOICE"""
voices = engine.getProperty('voices')       #getting details of current voice
#engine.setProperty('voice', voices[0].id)  #changing index, changes voices. o for male
engine.setProperty('voice', voices[1].id)   #changing index, changes voices. 1 for female


for filename in os.listdir('C://Users//Username//Desktop//E-books'):
    os.chdir('C://Users//Username//Desktop//E-books')
    path = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(path)
    Numberofpages = pdfReader.getNumPages()
    TotalText = ""

    for i in range(Numberofpages):
        page = pdfReader.getPage(i-1)
        text = page.extractText()
        TotalText = TotalText + text
        print("Parsing data:")
        print("Page {i} of {NumberOfPages}".format(i=i, NumberOfPages=Numberofpages))

    os.chdir('C://Users//Username//Downloads//Audiobooks')
    print("Making audiobook, please wait")
    engine.save_to_file(TotalText, str(filename)+'.mp3')
    engine.runAndWait()
    print("Done")
    time.sleep(5)
