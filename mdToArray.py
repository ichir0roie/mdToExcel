from myClasses import *

# maxSharp = 3

# indentSpaces = "  "

stng=Settings("settings.json")

class MdToArray:
    def __init__(self):
        self.md = ""
        self.book = Book()
        return

    def read(self, path: str):
        with open(path, encoding="utf-8", mode="r")as f:
            self.md = f.read()
        self.compile()
        return

    # HACK ちょっとこれは汚い｡
    def compile(self):
        self.book = Book()
        lines = self.md.split("\n")
        sheet = []
        # part,chapter,section,subsection
        pcss = [0, 0, 0, 0]
        sheetName = ""
        indent = 0
        for line in lines:
            column = []
            if len(line) <= 0:
                sheet.append([])
                continue
            # read # .
            p = 0
            while line[p] == "#":
                p += 1
            if p > 4:
                p = 4
                print("warning : deep indent. bad.")
            if p > 1:  # pcss
                indent = p-1-1
                pcss[indent] = pcss[indent]+1
                column.extend([""for i in range(indent)])
                column.append(str(pcss[indent])+".")
                column.append(line[p+1:])
                sheet.append(column)
                continue
            elif p == 1:  # sheet
                if sheetName != "" and len(sheet) > 0:
                    self.book.addSheet(sheetName, sheet)
                sheetName = line[p+1:]
                sheet = []
                pcss = [0, 0, 0, 0]
                continue
            # for list format
            if line[0] == " " or line[0] == "+" or line[1] == ".":
                listIndent = 0
                listTxt = ""
                sc = 0
                while line[sc] == " ":
                    sc += 1
            # for "+ "
                if line[sc] == "+":
                    listIndent = int(sc/stng.listPlusIndentSpaces)
                    addTex = line[line.find("+ ")+2:]
                # for "1. "
                elif line[sc+1] == ".":
                    listIndent = int(sc/stng.listNumIndentSpaces)
                    listTxt = line[sc:sc+2]
                    addTex = line[line.find(". ")+2:]
                column.extend([""for i in range(indent+1+listIndent)])
                column.append(listTxt)
                column.append(addTex)
            else:
                # generate nomal text
                column.extend([""for i in range(indent+1)])
                column.append(line[p:])
            sheet.append(column)
        
        if sheetName != "" and len(sheet) > 0:
            self.book.addSheet(sheetName, sheet)

    def printBook(self):
        for sheet in self.book.sheets:  # type:Sheet
            print("sheet : "+sheet.sheetName)
            # print(i)
            for line in sheet.data:
                t = ""
                for el in line:
                    if el == "":
                        el = "  "
                    t += el+" | "
                print(t)
            print("\n\n")


if __name__ == "__main__":
    mte = MdToArray()
    mte.read("mdDocs/sample.md")
    mte.compile()
    mte.printBook()
