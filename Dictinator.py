import sys

import PyPDF2
from docx import Document
from docx.shared import Pt


class Dictinator:

    def __init__(self,file,showASCIIFailure,filename):
        print(f"\nScanning {filename}:")
        self.fileReader = PyPDF2.PdfFileReader(file)
        self.pages = self.fileReader.pages
        self.words = self.getWordsWithPagenumber(self.pages)
        self.showASCIIFailure = showASCIIFailure

        self.sortWords()
        print(len(self.words), "words found")

        self.deleteTrash()
        print(len(self.words), "importand words left")

        self.outputText = self.createOutputText()
        print(len(self.outputText), "different words")



    def showDic(self):
        print(self.outputText)

    def returnWord(self,liste):
        return liste['word']

    def filterText(self, text):
        text = text.lower()

        text = text.replace("\n", "")
        text = text.replace("(", " ")
        text = text.replace(")", " ")
        text = text.replace("[", " ")
        text = text.replace("]", " ")
        text = text.replace("{", " ")
        text = text.replace("}", " ")

        text = text.replace("\"", "")
        text = text.replace(":", "")
        text = text.replace("=", " ")
        text = text.replace("“", "")
        text = text.replace("’", "")
        text = text.replace(",", "")
        text = text.replace("'", "")
        text = text.replace("%", " ")
        text = text.replace("/", " ")
        text = text.replace(".", " ")
        text = text.replace("|", " ")


        text = text.replace(" and ", " ")
        text = text.replace(" any ", " ")
        text = text.replace(" are ", " ")
        text = text.replace(" for ", " ")
        text = text.replace("hochschule esslingen", "")
        text = text.replace(" are ", " ")
        text = text.replace("−", " ")
        text = text.replace("", " ")
        text = text.replace("", " ")
        text = text.replace("-", " ")
        text = text.replace("•", " ")
        text = text.replace("~", " ")
        text = text.replace("�", " ")
        text = text.replace("  ", " ")
        text = text.replace("  ", " ")
        text = text.replace("  ", " ")
        text = text.replace("  ", " ")
        text = text.replace("  ", " ")
        text = text.replace("  ", " ")
        text = text.replace("  ", " ")

        return text

    def getWordsWithPagenumber(self,pages):
        number = 1
        paare = []


        for page in pages:

            filteredText = self.filterText(page.extract_text())
            for entry in filteredText.split():
                paare.append({"word": entry, "page": number})
            number += 1
        return paare

    def sortWords(self,):
        self.words.sort(key=self.returnWord)



    def deleteTrash(self,):
        index = 0
        trashList = []
        for entry in self.words:
            entry = entry["word"]
            if len(entry) < 3:
                trashList.append(index)

            index += 1

        trashList.reverse()
        for index in trashList:
            self.words.pop(index)

    def createOutputText(self):
        text = []
        lastAddedPageNumber = -1
        currentWord = str(self.words[0]["word"])
        pageListCurrentWord = currentWord

        for entry in self.words:

            #if the current word is the same as the one in the iteration, we add the page number to the List
            #Every word at the beginning in the line, followed by every apperance page number
            if entry["word"] == currentWord:

                #only add page number thats different then the number before
                if not lastAddedPageNumber == entry["page"]:
                    pageListCurrentWord += " S." + str(entry["page"])
                    lastAddedPageNumber = entry["page"]

            #if a new word is found we add the pageList to the final text
            #get The next Word and save it in currentWord
            #finaly we add the current word and oits page to the pageList
            else:
                try:
                    text.append(pageListCurrentWord)
                except:
                    pass
                currentWord = entry["word"]
                pageListCurrentWord = str(currentWord) + " S." + str(entry["page"])

                lastAddedPageNumber = -1


        return text

    def saveInFile(self,filename):
        file = open("Output/TXT/"+filename.split(".")[0]+'.txt', 'w')
        for row in self.outputText:
            try:
                file.write(row+"\n")
            except:
                if self.showASCIIFailure:
                    print("No ASCII:", row)

    def createWord(self,name):

        # create document
        doc = Document()
        table = doc.add_table(rows=int(len(self.outputText)/3)+1, cols=3,style="Table Grid")

        print(f"start creating Word from {name}")
        for index in range(len(self.outputText)):
            if index%3 == 0:
                hdr_cells = table.rows[int(index / 3)].cells
                self.update_progress(index/len(self.outputText))

            hdr_cells[index%3].text = self.outputText[index]

        doc.save("Output/word/"+name.split(".")[0]+".docx")


    def update_progress(self, progress):
        """
         Creates a progress bar
         :param progress: The progress thus far
         """
        barLength = 20
        status = ""
        if isinstance(progress, int):
            progress = float(progress)

        block = int(round(barLength * progress))
        text = "\r[{}] {:0.2f}% {}".format("#" * block + "-" * (barLength - block), progress * 100, status)
        sys.stdout.write(text)
        sys.stdout.flush()


