
# importing all the required modules
import Dictinator
from os import listdir

# creating an object
if __name__== "__main__":
    showASCIIFailure = False
    for filename in listdir("./input"):
        PDF = open(f'input/{filename}', 'rb')
        Dic = Dictinator.Dictinator(PDF,showASCIIFailure,filename)
        Dic.saveInFile(filename)
        Dic.createWord(filename)

    print("Ready for Exam!\nGood Luck")


