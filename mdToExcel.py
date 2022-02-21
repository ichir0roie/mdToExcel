

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

        self.readFolder()


    # まずエクセル共をコピーして貼り付けておく｡
    def copyExcels(self):
        targets=[p for p in self.targetPaths if not ".md" in p]

        for path in targets:
            self.save(path)

        return

    def readFolder(self):
        # 対象のフォルダ内のファイルを全部読み込んで、リストに格納

        searchPath=(self.stng.inputMdFolder+"\\**\\**.*").replace("\\\\", "\\")
        self.targetPaths=[p for p in glob.glob(searchPath,recursive=True) if os.path.isfile(p)]
        self.targetPaths=[p for p in self.targetPaths if not ".git" in p]
        if not test:
            self.targetPaths=[p for p in self.targetPaths if not "test" in p]
        else:
            self.targetPaths=[p for p in self.targetPaths if "test" in p]
        return

    def generate(self):
        # 対象のデータを順次変換

        targets=[p for p in self.targetPaths if ".md" in p]

        for path in targets:
            self.mta.read(path)
            self.ate.setBook(self.mta.book)
            savePath=path.replace(self.stng.inputMdFolder, self.stng.outputExFolder)
            extension=self.stng.templateBookPath[self.stng.templateBookPath.find("."):]
            savePath=savePath.replace(".md", extension)
            self.ate.generateBook(savePath,self.stng.font,self.stng.size)

    def save(self,path):
        savePath=path.replace(self.stng.inputMdFolder, self.stng.outputExFolder)
        saveDir=savePath.replace("\\", "/")
        saveDir=saveDir[:saveDir.rfind("/")]
        if not os.path.exists(saveDir):
            os.makedirs(saveDir)
        shutil.copy(path,savePath)

if __name__ == "__main__":
    # test=True

    m = MdToExcel()
    m.copyExcels()
    m.generate()


