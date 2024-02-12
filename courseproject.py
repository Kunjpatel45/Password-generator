import openpyxl
import pyinputplus as pyip
import random
import re

# It will a regular expression that matches the alphabetical order
pattern=r'^[a-zA-Z]+$'
regexObject = re.compile(pattern)

def getValidName():
    #This function wil ask user to enter the name(only alphabetical order);it will give upto 3 chance to user to enter valid name
    for i in range(3):  # allow the user to enter the name up to 3 times
        name = input("Can you please enter your name: ")
        if regexObject.match(name):
            return name
        else:
            print("If possible, enter the name using alphabet's only; Don't use any other type like number or any special character.")
    print("Sorry! You have used all the attempts")

def getValidPassword():
    #This function wil ask user to enter the length of the password and it will give upto 3 chance to user to enter valid length
    for i in range(3):
        length = pyip.inputInt("enter the length of password (It must be of 8 character or greater than that): ")
        if length < 8:
            print("Error: Password must be at least 8 characters long.")
        else:
            password = generatePassword(length)
            return password
    print("Sorry! You have used all the attempts")

def generatePassword(length):
    # This function will generate a password of the length that user enters in above function
    upperLetters =["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    lowerLetters =["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"]
    number = ["1","2","3","4","5","6","7","8","9","0"]
    specialChar = ["'","!","@","#","$","%","^","&","*","(",")","_","-","+","=","{","}","[","]","|",":",";","<",">",",",".","?","/","~","`"]
    allChars = upperLetters + lowerLetters + number + specialChar
    password = ''.join(random.sample(allChars, length))
    return password

def appendToExcelFile(name, password):
    #Append the name and password to the 'project.xlsx' Excel file.
    workbook = openpyxl.load_workbook('project.xlsx')
    sheet = workbook.active
    row = (name, password)
    sheet.append(row)
    workbook.save('project.xlsx')
    print("Your Password is:",password)
    print("Your name and password has been saved to an excel sheet named 'project.xlsx'.")

def main():
    #Main function.
    name = getValidName()
    password = getValidPassword()
    appendToExcelFile(name, password)

main()
