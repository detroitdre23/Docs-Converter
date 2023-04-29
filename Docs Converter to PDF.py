import sys
import os
import win32com.client
from time import sleep

def word(input_folder_path):
    input_file_paths = os.listdir(input_folder_path)
    
    for input_file_name in input_file_paths:
    
        if not input_file_name.lower().endswith((".docx",".doc")):
            continue
            
        input_file_path = os.path.join(input_folder_path, input_file_name)
            
        word = win32com.client.Dispatch("Word.Application")
        
        docs = word.Documents.Open(input_file_path)
        
        word.Visible = False
        
        file_name = os.path.splitext(input_file_name)[0]
        
        output_file_path = os.path.join(output_folder_path, file_name + ".pdf")
        
        docs.SaveAs(output_file_path, 17)
        
        docs.Close()
        
        os.remove(input_file_path)
        
        print("file ",input_file_name," converted successfully")
        print("file ",input_file_name," deleted successfully")
        
        
  
def powerpoint(input_folder_path):
    input_file_paths = os.listdir(input_folder_path)

    for input_file_name in input_file_paths:

        if not input_file_name.lower().endswith((".ppt", ".pptx", ".ppsx")):
            continue
        
        input_file_path = os.path.join(input_folder_path, input_file_name)
            
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
       
        slides = powerpoint.Presentations.Open(input_file_path,WithWindow=False)
        
        file_name = os.path.splitext(input_file_name)[0]
        
        output_file_path = os.path.join(output_folder_path, file_name + ".pdf")
       
        slides.SaveAs(output_file_path, 32)
        
        slides.Close()
        
        os.remove(input_file_path)
        
        print("file ",input_file_name," converted successfully")
        print("file ",input_file_name," deleted successfully")
        
        os.system("TASKKILL /F /IM powerpnt.exe >nul 2>&1")    
    
    
if __name__ == "__main__":
    os.system("cls")
    print("Welcome To Documents/Slides Converter to pdf\n")
    print("1 - Word To PDF")
    print("2 - Powerpoint To PDF")
    print("Choose an option")
    choice = input()
    choice = int(choice)
    if(choice==1):
        print("please enter input folder : ")
        input_folder_path=input()
        print("please enter output folder : ")
        output_folder_path=input()
        input_folder_path = os.path.abspath(input_folder_path)
        output_folder_path = os.path.abspath(output_folder_path)
        word(input_folder_path)
        sleep(5)
        exit(0)
    elif(choice==2):
        print("please enter input folder : ")
        input_folder_path=input()
        print("please enter output folder : ")
        output_folder_path=input()
        input_folder_path = os.path.abspath(input_folder_path)
        output_folder_path = os.path.abspath(output_folder_path)
        powerpoint(input_folder_path)
        sleep(5)
        exit(0)
    else:
        print("error")
        
    