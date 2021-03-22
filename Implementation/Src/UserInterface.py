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
Creating class with name UserInterface which is calling itself
'''
class UserInterface:

    # def __init__(self, UserInterface):
    #     self.UserInterface  = UserInterface

    def textfilenameread(self):
        pass

        '''
        This fuction textfilepathread()  is taking the text file path from user and then copying all the work book path
        in the globla pathlist so to available for every function.
        '''

    def textfilepathread(self):

        textfilepath = input("Enter your .txt Path here:  ")
        file = open(textfilepath, "r")
        context = file.read()
        textpathlist = context.split("\n")  # this will seperate all teh path stored in text file.
            # print(textpathlist)
        Dict["RF_TextfilePath"] = 1
        for item in textpathlist:
            pathList.append(item)  # appending the path one by one to the pathlist.

        '''
        Function terminalpathinpt() taking the all the file path from the terminal at once.
        when the user will press enter key twice then the program stored all the path in the pathlist ariable

        '''

    def terminalpathinput(self):
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

    def manuentrypath(self):
        manuInput = input("Copy Your Path Here : ")
        pathList.append(manuInput)  # appending the path list
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

    def manuentryname(self):
        pass

        '''

        fucntion name :   user_selection()
        This function invoke first as program started. the function give the choice to the user for multiple 
        user selection option and to nvigate to other input method.

        '''
    def user_selection(self):

        print("Enter the choice given below. Use Number for your selected choice")
        print("1. Add Text File  Name")
        print("2. Give path of .txt which contain Workbook path ")
        print("3. Copy All the Workbook Path directly to terminal")
        print("4. Manually You want to give the Workbook Path")
        print("5. Manually Giving the Workbook Name Present in the same directory")
        print("6. Exit ")
        choice = int(input("Waiting for your choice : "))
        if choice == 1:
            UserInterface.textfilenameread(self)
        elif choice == 2:
            UserInterface.textfilepathread(self)
        elif choice == 3:
            UserInterface.terminalpathinput(self)
        elif choice == 4:
            UserInterface.manuentrypath(self)
        elif choice == 5:
            UserInterface.manuentryname(self)
        elif choice == 6:
            return

'''
Creating another class with UserChoiceSlection and keeping all the function inside them . the class is calling self.
'''
class UserChoiceSelection:

    def manual_filter_input(self):

        print("Enter Maximum Three Identity :: Format :: Col_Name,value ::: Ex- xyz,123abc :::: PRESS Enter key")
        filterinput = input("Input Data Here : ")
        # print(filterinput)
        inputlist1 = filterinput.split(
            ",")  # Here taking input in list but filtering the data according to second data of list.
        inputList.append(inputlist1[1])
        # inputDict[inputlist1[0]] = inputlist1[1]
        Dict["RF_ManualFilterInput"] = 1  # raising the flag
        userInput = input("Want to add more filter:: Y/N : ")
        while True:
            if userInput == 'Y' or userInput == 'y':
                filterinput = input("Input Data Here : ")
                inputlist1 = filterinput.split(",")  # spliting the choice
                inputList.append(inputlist1[1])
                # inputDict[inputlist1[0]] = inputlist1[1]

            elif userInput == 'N' or userInput == 'n':
                break
            userInput = input("Want to add more filter:: Y/N : ")

    '''

    Funtion name: show_me_data()
    The function taking the chouice for showing the data by iput key search or by automatic search.

    '''


    def show_me_data(self):
        print("1. Search By input")
        print("2. Automatic Search")
        choice = int(input(" Enter Choice"))

        if choice == 1:
            UserChoiceSelection.manual_filter_input(self)
            Dict["Flag_showData"] = 1
        elif choice == 2:
            Dict["Flag_showData"] = 0

    '''
    '''
    def choice_selection(self):
        print(" 1. Filter the Data according to your choice ")
        print(" 2. Want to search filter input ")
        userChoice = int(input("Enter Your choice here : "))
        if userChoice == 1:
            UserChoiceSelection.manual_filter_input(self)
        elif userChoice == 2:
            UserChoiceSelection.show_me_data(self)


class InputOutput(UserChoiceSelection):       # this class inheriting the UserChoiceSelection class and calling self.

    def create_workbook(self):
        book = Workbook()
        sheet = book.active
        sheet.title = "Sheet1"
        book.save(outputPath[0])

    """
    Function Name : desktop_output()

    this function is creating the output file by default on desktop
    it takes only output file name and using the os library to do the same automatic.

    """

    def desktop_output(self):
        import os
        name = input("Enter New Work Book Name: ")
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        newpath = desktop + "\\" + name + ".xlsx"
        outputPath.append(newpath)
        InputOutput.create_workbook(self)

    """

    Function name : give_path()
    Input : user output store data folder path
    Output : appending the path in the outputpath list

    """

    def give_path(self):

        choice = input(" Paste Your Path Here : ")
        outputPath.append(choice)

    """

    Function name : outputResult_path()
    Input :  Taking the choice to store the data from user ///// Taking path from user
    Output :  calling the respective function after getting the user choice.


    """
    """

    """

    def outputResult_path(self):
        print("\n")
        print(" Give Choice For Output Result")
        print("1. Create New WorkBook on Desktop")
        print("2. Give Existing Work Book Path")
        choice = int(input("Enter Your Choice Here : "))
        if choice == 1:
            InputOutput.desktop_output(self)
        elif choice == 2:
            InputOutput.give_path(self)


# obj1 = UserInterface()  # creating object for first class
# obj2 = InputOutput()    # creating object for second class
# obj1.user_selection()    # invoking the function inside the class for user selection
# obj2.choice_selection()  # invoking the function inside the class
# obj2.outputResult_path()
# print(pathList)
# print(inputList)
# print(outputPath)





