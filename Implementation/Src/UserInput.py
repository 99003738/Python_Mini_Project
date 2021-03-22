'''

Importing the library Openpyxl
'''
from openpyxl import Workbook
import time
'''
Decalring the Global variable
'''
pathList = []     # TO STORE THE PATH GIVEN BY USER
inputList = []    # To Store the search key given by user
'''
This is flag to indicate which type of input came from the user.
'''
Dict = dict({"RF_TextFilePath": 0, "RF_ManualPath" : 0, "RF_ManualFilterInput": 0, "Flag_showData": 0})
inputDict = {}
'''
this variable is used to store the output path given by user
'''
outputPath = []
'''
Function to take the input from text file name.
Featured currently not implemented
'''
def textfilenameread():
    pass

'''
This fuction textfilepathread()  is taking the text file path from user and then copying all the work book path
in the globla pathlist so to available for every function.
'''
def textfilepathread():

    textfilepath = input("Enter your .txt Path here:  ")
    file = open(textfilepath, "r")
    context = file.read()
    textpathlist = context.split("\n")     # this will seperate all teh path stored in text file.
    # print(textpathlist)
    Dict["RF_TextfilePath"] = 1
    for item in textpathlist:
        pathList.append(item)               # appending the path one by one to the pathlist.

'''
Function terminalpathinpt() taking the all the file path from the terminal at once.
when the user will press enter key twice then the program stored all the path in the pathlist ariable

'''


def terminalpathinput():
    print("Copy all the below and hit double Enter key :")
    while True:  # directly copy all the path in terminal with new line and hit enter and enter
        line = input()  # the directory file will loaded into the list form and can be extract by list index
        if line:
            pathList.append(line)
        else:
            break


'''

menualentrypath() ---  function take first path and store it into the pathlist and then ask for user
 weather he wants to add one more path. if user says yes Y then it will again take input
 and store it into the respective variable.
 the program get terminated ehen the user input the NO N keyword.
 '''


def manuentrypath():
    manuInput = input("Copy Your Path Here : ")
    pathList.append(manuInput)          # appending the path list
    Dict["RF_ManualPath"] = 1
    userInput = input("Want to Add more:: Y/N : ")
    while True:
        if userInput == 'Y' or userInput == 'y':
            manuInput = input("Copy : ")
            pathList.append(manuInput)
        elif userInput == 'N' or userInput == 'n':
            break
        userInput = input("Want to Add more:: Y/N : ")

'''

This function is currently not used . kept for future development

'''
def manuentryname():
    pass

'''

fucntion name :   user_selection()
This function invoke first as program started. the function give the choice to the user for multiple 
user selection option and to nvigate to other input method.

'''



def user_selection():

    print("Enter the choice given below. Use Number for your selected choice")
    print("1. Add Text File  Name")
    print("2. Give path of .txt which contain Workbook path ")
    print("3. Copy All the Workbook Path directly to terminal")
    print("4. Manually You want to give the Workbook Path")
    print("5. Manually Giving the Workbook Name Present in the same directory")
    print("6. Exit ")
    choice = int(input("Waiting for your choice : "))
    if choice == 1:
        textfilenameread()
    elif choice == 2:
        textfilepathread()
    elif choice == 3:
        terminalpathinput()
    elif choice == 4:
        manuentrypath()
    elif choice == 5:
        manuentryname()
    elif choice == 6:
        return

'''

Funtion name : manual_filter_input()
Asking user to enter the searh key to which he wants to take out data from the Workbook

'''
def manual_filter_input():

    print("Enter Maximum Three Identity :: Format :: Col_Name,value ::: Ex- xyz,123abc :::: PRESS Enter key")
    filterinput = input("Input Data Here : ")
    # print(filterinput)
    inputlist1 = filterinput.split(",")                              # Here taking input in list but filtering the data according to second data of list.
    inputList.append(inputlist1[1])
    # inputDict[inputlist1[0]] = inputlist1[1]
    Dict["RF_ManualFilterInput"] = 1                   # raising the flag
    userInput = input("Want to add more filter:: Y/N : ")
    while True:
        if userInput == 'Y' or userInput == 'y':
            filterinput = input("Input Data Here : ")
            inputlist1 = filterinput.split(",")         # spliting the choice
            inputList.append(inputlist1[1])
            # inputDict[inputlist1[0]] = inputlist1[1]

        elif userInput == 'N' or userInput == 'n':
            break
        userInput = input("Want to add more filter:: Y/N : ")


'''

Funtion name: show_me_data()
The function taking the chouice for showing the data by iput key search or by automatic search.

'''
def show_me_data():
    print("1. Search By input")
    print("2. Automatic Search")
    choice = int(input(" Enter Choice"))

    if choice == 1:
        manual_filter_input()
        Dict["Flag_showData"] = 1
    elif choice ==2:
        Dict["Flag_showData"] = 0

'''
'''

def user_choice_selection():
    print(" 1. Filter the Data according to your choice ")
    print(" 2. Want to search filter input ")
    userChoice = int(input("Enter Your choice here : "))
    if userChoice == 1:
        manual_filter_input()
    elif userChoice == 2:
        show_me_data()


def create_workbook():
    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    book.save(outputPath[0])

"""
Function Name : desktop_output()

this function is creating the output file by default on desktop
it takes only output file name and using the os library to do the same automatic.

"""



def desktop_output():
    import os
    name = input("Enter New Work Book Name: ")
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    newpath = desktop + "\\" + name + ".xlsx"
    outputPath.append(newpath)
    create_workbook()

"""

Function name : give_path()
Input : user output store data folder path
Output : appending the path in the outputpath list

"""
def give_path():

    choice = input(" Paste Your Path Here : ")
    outputPath.append(choice)
"""

Function name : outputResult_path()
Input :  Taking the choice to store the data from user ///// Taking path from user
Output :  calling the respective function after getting the user choice.

    
"""
"""

"""
def outputResult_path():
    print("\n")
    print(" Give Choice For Output Result")
    print("1. Create New WorkBook on Desktop")
    print("2. Give Existing Work Book Path")
    choice = int(input("Enter Your Choice Here : "))
    if choice == 1:
        desktop_output()
    elif choice == 2:
        give_path()

# user_selection()
# user_choice_selection()
# outputResult_path()

# print(pathList)



