

import mdToArray
import arrayToExcel
import json

import glob

from myClasses import *

# import settings as self.stng

import os
import shutil



class MdToExcel:
    def __init__(self):
        self.mta = mdToArray.MdToArray()
        self.ate = arrayToExcel.ArrayToExcel()
        self.stng=Settings("settings.json")
        
        self.stng.inputMdFolder=self.stng.inputMdFolder.replace("/", "\\")
        self.stng.outputExFolder=self.stng.outputExFolder.replace("/", "\\")

        # 追記上書きにしたい｡
        # shutil.rmtree(self.stng.outputExFolder)
        # os.mkdir(self.stng.outputExFolder)

        self.ate.readTemplate(self.stng.templateBookPath,self.stng.templateSheetName, self.stng.startRow, self.stng.startCol)

        # with open("settings.json", mode="r", encoding="utf-8")as f:
        #     text = f.read()
        #     self.settings = json.loads(text)

        self.targetPaths = []


    def readFolder(self):
        # 対象のフォルダ内のファイルを全部読み込んで、リストに格納

        searchPath=(self.stng.inputMdFolder+"\\**\\**.md").replace("\\\\", "\\")
        self.targetPaths=glob.glob(searchPath,recursive=True)
        print(self.targetPaths)
        return

    def generate(self):
        # 対象のデータを順次変換

        for path in self.targetPaths:
            self.mta.read(path)
            self.ate.setBook(self.mta.book)
            savePath=path.replace(self.stng.inputMdFolder, self.stng.outputExFolder)
            extension=self.stng.templateBookPath[self.stng.templateBookPath.find("."):]
            savePath=savePath.replace(".md", extension)
            self.ate.generateBook(savePath,self.stng.font,self.stng.size)


if __name__ == "__main__":
    m = MdToExcel()
    # print(m.settings)
    m.readFolder()
    m.generate()


