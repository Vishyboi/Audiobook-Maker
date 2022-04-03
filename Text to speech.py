#This is a program to make an audiobook version of every file in the directory
#Run the makeAudiobook function with appropriate arguements, and wait for the program to finish.
#It can make pdf, docs, txt files to mp3
#In case you want any other extension, define a function to get its text, and then make a IsExtension variable, and an if statement
#to check if the file is of that extension, and then call the function to make the audio file.
#Done by Vishyboi123
#Version 1.0

import pyttsx3, sys, os, PyPDF2, time, docx, pyinputplus
from pathlib import Path

def docText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)


engine = pyttsx3.init() # object creation

def getTextDocx(Filename):
    doc = docx.Document(Filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    TotalText = ''
    for i,j in enumerate(fullText):
        TotalText = TotalText+str(j)
    return TotalText

def getTextPdf(Filename):
    path = open(Filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(path)
    Numberofpages = pdfReader.getNumPages()
    TotalText = ""

    for i in range(Numberofpages):
        page = pdfReader.getPage(i-1)
        text = page.extractText()
        TotalText = TotalText + text
        print("Parsing data")
        print("Page {i} of {NumberOfPages} from {filename}".format(i=i, NumberOfPages=Numberofpages,filename=filename))
        return TotalText

def getTextTxt(Filename):
    print("Parsing data")
    path=open(Filename, 'r')
    Totaltext = path.read()
    return Totaltext

def MakeAudio(filename, Text, TargetDir):
    print("Initializing making the audio file")
    os.chdir(TargetDir)
    engine.save_to_file(TotalText, str(filename)+'.mp3')
    engine.runAndWait()
    print("Done")
    time.sleep(5)


#Arguements = Directory of files, The extension of which type of file you want to make as an audiobook
#Optional arguements - word per minute speed of speaker, volume and person speaking it
def makeAudiobook(directory, wpm=175, vol=1.0, speaker=male, targetdir):
    float(vol)
    volume = engine.getProperty('volume')   #getting to know current volume level (min=0 and max=1)
    engine.setProperty('volume',vol) 

    rate = engine.getProperty('rate')   # getting details of current speaking rate
    engine.setProperty('rate', wpm)     # setting up new voice rate

    voices = engine.getProperty('voices')       #getting details of current voice
    if speaker is male:
        engine.setProperty('voice', voices[0].id)
    else:
        engine.setProperty('voice', voices[1].id)

    try:
        sys.path.append(directory) # To include directory of the files
    except FileNotFoundError:
        print("Directory not found, please try again")
        sys.exit()
    
    Path(directory)
    IsDirectory = os.path.isdir(directory)
    Path(targetdir)
    FinalDirectory = os.path.isdir(targetdir)
    
    if IsDirectory is False or FinalDirectory is False:
        print("Directory does not exist")
        sys.exit()

    else:

        Path(directory)
        for filename in os.listdir(directory):
            os.chdir(directory)
            IsPDF = filename.endswith(.pdf)
            IsDOC = filename.endswith(.doc|.docx)
            IsTXT = filename.endswith(.txt)
    
            if IsPDF is True:
                FinalText = getTextPdf(filename)
                MakeAudio(filename, FinalText, TargetDir)

            elif IsDOC is True:
                FinalText = getTextDocx(filename)
                MakeAudio(filename, FinalText, TargetDir)

            elif IsTXT is True:
                FinalText = getTextTxt(filename)
                MakeAudio(filename, FinalText, TargetDir)

            else:
                print("Cannot read {filename}".format(filename=filename)))


def SingleFileAudio(file, fileDir, Targetfiledir, wpm=175, vol=1.0, speaker=male):
    float(vol)
    volume = engine.getProperty('volume')   #getting to know current volume level (min=0 and max=1)
    engine.setProperty('volume',vol) 

    rate = engine.getProperty('rate')   # getting details of current speaking rate
    engine.setProperty('rate', wpm)     # setting up new voice rate

    voices = engine.getProperty('voices')       #getting details of current voice
    if speaker is male:
        engine.setProperty('voice', voices[0].id)
    else:
        engine.setProperty('voice', voices[1].id)

    Path(fileDir)
    Path(Targetfiledir)

    try:
        sys.path.append(fileDir) # To include directory of the files
    except FileNotFoundError:
        print("Directory not found, please try again")
        sys.exit()
    
    Path(directory)
    IsDirectory = os.path.isdir(directory)
    Path(targetdir)
    FinalDirectory = os.path.isdir(targetdir)
    
    if IsDirectory is False or FinalDirectory is False:
        print("Directory does not exist")
        sys.exit()

    else:
        if file.endswith(.pdf):
            FinalText = getTextPdf(file)
            MakeAudio(file, FinalText, Targetfiledir)

        elif file.endswith(.doc|.docx):
            FinalText = getTextDocx(file)
            MakeAudio(file, FinalText, Targetfiledir)

        elif file.endswith(.txt):
            FinalText = getTextTxt(file)
            MakeAudio(file, FinalText, Targetfiledir)


while True:
    print("Welcome to the AudioBook Maker")
    print("Please select an option")
    inp = pyinputplus.inputMenu(['Make an Audiobook of every file in a directory', 'Make an Audio File of a single file', 'Quit'])
    if inp is 'Make an Audiobook of every file in a directory':
        print("Please enter the directory of the files")
        directory = pyinputplus.inputStr(prompt="Enter the directory of the files", allowRegexes=['.*'])
        print("Please enter the directory you want to save the audiobook")
        targetdirectory = pyinputplus.inputStr(prompt="Enter the directory you want to save the audiobook", allowRegexes=['.*'])
        print("Please enter the speed of the speaker")
        wordpm = pyinputplus.inputInt(prompt="Enter the speed of the speaker", min=1, max=500)
        print("Please enter the volume of the speaker")
        vol = float(pyinputplus.inputFloat(prompt="Enter the volume of the speaker", min=0, max=1))
        print("Please select the speaker (male or female)")
        speakerPerson = pyinputplus.inputMenu(['Male','Female'])

        makeAudiobook(directory, wpm=wordpm, volume=vol, speaker=speakerPerson, targetdirectory)
        continue

    
    elif inp is 'Make an Audio File of a single file':
        print("Please enter the directory of the files")
        directory = pyinputplus.inputStr(prompt="Enter the directory of the files", allowRegexes=['.*'])
        print("Please enter the file you want to make an audiobook")
        filename = pyinputplus.inputStr(prompt="Enter the file you want to make an audiobook", allowRegexes=['.*'])
        print("Please enter the directory you want to save the audiobook")
        targetdirectory = pyinputplus.inputStr(prompt="Enter the directory you want to save the audiobook", allowRegexes=['.*'])
        print("Please enter the speed of the speaker")
        wordpm = pyinputplus.inputInt(prompt="Enter the speed of the speaker", min=1, max=500)
        print("Please enter the volume of the speaker")
        volume = float(pyinputplus.inputFloat(prompt="Enter the volume of the speaker", min=0, max=1))
        print("Please select the speaker (male or female)")
        speakerPerson = pyinputplus.inputMenu(['Male','Female'])

        SingleFileAudio(filename, directory, targetdirectory, wpm=wordpm, vol=volume, speaker=speakerPerson)
        continue

    elif inp is 'Quit':
        print('Thank you for using audiobook maker')
        sys.exit()