#! encoding: utf-8

__Author__ = 'Yang'

#引包
import os.path
import re
import win32com
from win32com.client import Dispatch, constants

global filelist

def readfile(fileabspath):
    file_object = open(fileabspath)
    try:
        all_the_text = file_object.read()
        print fileabspath
    finally:
        file_object.close()
        return all_the_text
    
def processDirectory(args, dirname, filenames):

    for filename in filenames:
        if re.search(r"[^\s]\.java", filename) != None:
            filelist[0].append(dirname + '\\' + filename)
##        elif re.search(r"[^\s]\.c", filename) != None:
##            filelist[1].append(dirname + '\\' + filename)
##        elif re.search(r"[^\s]\.cpp", filename) != None:
##            filelist[2].append(dirname + '\\' + filename)



if __name__ == '__main__':
    filelist = [[], [], []]
    os.path.walk(r'E:\sources\python\file_process\MyGraduationProject', processDirectory, None)
    
    w = win32com.client.Dispatch('Word.Application')
    w.Visible = 0
    w.DisplayAlerts = 0

##    doc = w.Documents.Add()
    doc = w.Documents.Open(FileName = r'E:\sources\python\file_process\text.docx')
    

    wrange = doc.Range()
    for ele in filelist:
        for hfile in ele:
            wrange.InsertAfter('// ' + hfile + '\n')
            wrange.InsertAfter(readfile(hfile) + '\n')
    


    #保存doc文件(为.docx后缀)
    ##    doc.SaveAs(r'E:\sources\python\file_process\Passthru\text.docx')
    doc.Save()
    w.Documents.Close()
    w.Quit()
    
        
