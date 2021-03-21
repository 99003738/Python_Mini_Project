# Problem Statement

Write a program which takes user input and search all the data present in the given workbook of excel sheets and showing the data in a excel sheet named with Master sheet present in the same workbook.

# Flow of Problem statement

## Step01 : Taking User Input
            Maximum three different data allowed to take from the user. 
            The data could be anything based on the work book in which the search operation going to perform.
            Example:
                    Name : XYZ
                    PS Number: 1234XXX
                    Email : abc12@mail.com
                    
## Step02 : Validating the incoming user input.
            •	Checking whether the input data is correct or not. 
              This task can be achieved by going through the same workbook and traversing along all the sheets and finding out                
              the corresponding data and tally with the input data.
            Example :
                      Name : XYZ
                      PS Number : 1234XXX
                      Email : abc12@mail.com
           Checking the data in all sheet present in same workbook.
           If found authentic then throwing message “Data matched with given Input”
           If not found authentic then throwing message “Invalid data : No match found”
           
## Step03 : Searching the data with input and saving in different sheet in the same workbook with name “Master’” sheet.
            •	Going through all the sheets one by one and saving the matched data in Master sheet. 
            •	The matched data can be appended row wise or column wise according to the user choice.
            
## Step04 : Storing data in Master sheet
            •	The matched data will be stored in the Master sheet with following structure explained below:
                  i.	The first row should be merged with all the columns.
                  
                 ii.	Print the data in first row as given format:
                       PS Number : 1234XXXX, Name : XYZ, Email : abc12@mail.com
                       
                iii.	Second row contain the header of the sheet data which is going to be saved in master sheet.
                
                 iv.	If row appended data is chosen as output sheet structure format then leave a space of one row and then
                        print header of the second data sheet followed by the data of corresponding sheet.
                      
                 v.	  Constraint : Don’t  print those  column of any data sheet which contain PS Number, Name and Email.
                  
                vi.  	If Column appended data is chosen as output sheet structure format then follow the same procedure for row 1 but row 2 contain the 
                      header of matched data of corresponding data sheet followed by other sheet header if any matched found.
                      
               vii.	  Row appended and column appended choice depend upon the sheet header format weather the workbook containing
                        the same set of header or different header name.

# 4W & 1H :

### What:
             • We are having multiple excel workbook stored in diffenrent location in our system, to which we need to extract all the worksheet in them.
             • Every Work Book have different number of work sheet in them.
             • Taking a input search key and finding out all the data across the work book.
             • It is used for easy search of a particular cell or data of a person
             • It provides information of every person details which is present inside the loaded work book.
### When:
             • Searching all the data of a given search key.
             • To get the contact information.
             • To get the required details of that person according to data formulated in the sheet.
### Why:
             • We are using to retrieve tha data of an individual candidate from the excel workbook of 5 sheets where all the relevant data of 40 candidates is present.
             • We can easily access the details of that individual by giving some input such as name , ps no and email id.
### Where:
             • To check the information and bio of a person
             • Very useful during emergency times like health issues
             • We can also use it for knowing that person's bank details and other details related to his or her educational qualification.
### How:
             • Input:- We need to give maximum  3 inputs such as Name, Ps No and Email Id. 
             • The Input is dyanmic means either one or two or three can be given as a search key.
             • Output: -We will get all the relevant information of that person whose name, ps no and email id is given.
             • source: -All the relevant data will get copied in master sheet created by the user choice.

# SWOT Analysis

![abhi](https://user-images.githubusercontent.com/78892310/111924442-9b644a00-8aca-11eb-9cc2-6d9053a9e4db.PNG)
