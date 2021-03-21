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

