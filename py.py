import win32com
from genericpath import exists
from win32com.client import Dispatch, constants
import sys,os 

def ForSpecificFileInRootDir(process_function,file_type='.*'):
    items = os.listdir(".")
    file_list = []
    for names in items:
        if names.endswith(file_type) and names[0]!='~':
           file_list.append(names)

    print(file_list)
    for file in file_list:
        process_function(file)

def ExtractWords(ppt_file):
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    ppt.Visible = 1
    local_path=sys.path[0]
    pptSel = ppt.Presentations.Open(local_path+'\\'+ppt_file)
    if exists(local_path+'\\output.txt'):
        f = open(local_path+'\\output.txt','a',encoding='utf-8')
    else:
        f=open(local_path+'\\output.txt','w',encoding='utf-8')
    slide_count = pptSel.Slides.Count
    for i in range(1,slide_count + 1):
        shape_count = pptSel.Slides(i).Shapes.Count
        print(shape_count)
        for j in range(1,shape_count + 1):
            if pptSel.Slides(i).Shapes(j).HasTextFrame:
                s=pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
                if s!='':
                   f.write(s+'\n')
    f.close()
    pptSel.Close()

ForSpecificFileInRootDir(ExtractWords,'.ppt')

