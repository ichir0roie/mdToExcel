
import os
import json

class Book:
    def __init__(self):
        self.sheets = []

    def addSheet(self, name: str, sheetArray: list):
        sheet = Sheet(name, sheetArray)
        self.sheets.append(sheet)

class Sheet:
    def __init__(self, name, data):
        self.sheetName = name
        self.data = data


class Settings:
    def __init__(self,path:str):
        self.templateBookPath=None
        self.templateSheetName=None
        self.inputMdFolder=None
        self.outputExFolder=None
        self.combineFolder=None
        self.ignoreSheetNames=None
        self.startRow=None
        self.startCol=None
        self.font=None
        self.size=None

        self.extracts=[]

        if os.path.exists(path):
            with open(path,mode="r",encoding="utf-8")as f:
                text=f.read()
                self.stng=json.loads(text)

            self.templateBookPath=self.stng["templateBookPath"]
            self.templateSheetName=self.stng["templateSheetName"]
            self.inputMdFolder=self.stng["inputMdFolder"]
            self.outputExFolder=self.stng["outputExFolder"]
            self.combineFolder=self.stng["combineFolder"]
            self.ignoreSheetNames=self.stng["ignoreSheetNames"]
            self.startRow=self.stng["startRow"]
            self.startCol=self.stng["startCol"]
            self.font=self.stng["font"]
            self.size=self.stng["size"]

            self.listPlusIndentSpaces=self.stng["listPlusIndentSpaces"]
            self.listNumIndentSpaces=self.stng["listNumIndentSpaces"]
            self.extracts=self.stng["extracts"]


test=False