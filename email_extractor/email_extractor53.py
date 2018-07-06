import re
import os
import csv
import win32com.client
import datetime
import Tkinter as tk
import tkFileDialog
#---------- constants ----------#
#create RegEx for emails (what counts as an email)
global emailRegEx
#emailRegEx = re.compile(r'\w\S+@\w+\.\S+\w+')
emailRegEx = re.compile(r'[a-zA-Z0-9\.\_\-]+@\w+\.\S+\w+')
#create outlook object
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace('MAPI')
#name of output file
outFile = 'emails_' + str(datetime.date.today()) + '.csv'

class Body(tk.Frame):
    def __init__(self, master, buttons):
        #labels that will appear on the GUI body, with variable text
        currentDirLabel = tk.Label(master, textvariable = buttons.selectedDir)
        currentDirLabel.grid(row=1,column=1, columnspan=3)
        statusLabel = tk.Label(master, textvariable= buttons.statusText)
        statusLabel.grid(row=2,column=2)
        
class Buttons(tk.Frame):
    def __init__(self, master):
        changeDirButton = tk.Button(master, text='Change Directory...',
                                    command= lambda: self.update_directory(master, self.select_directory(master),self.selectedDir))
        changeDirButton.grid(row=1, column=0, sticky=tk.W+tk.E)
        self.folder_var = tk.StringVar()
        self.folder_var.set("Select Folder")        
        selectFolderMenu = tk.OptionMenu(master, self.folder_var, *(getFolderList(outlook)[1]))
        selectFolderMenu.grid(row=2,column=0)        
        extractButton = tk.Button(master, text="Extract",
                                  command = lambda: self.update_status(master, self.statusText, self.folder_var.get(), getFolderList(outlook), getFolderDic(outlook)))
        extractButton.grid(row=4,column=4, sticky=tk.W+tk.E)
        
        self.selectedDir = tk.StringVar()
        self.selectedDir.set("Working Directory: " + os.getcwd())
        self.statusText = tk.StringVar()
        self.statusText.set("Select mail folder and hit extract!")

    def select_directory(self, master):
        directory = tkFileDialog.askdirectory(title='Select folder datafile is located in...')
        try:
            os.chdir(directory)
            print("Directory has been changed to " + directory)
            return directory
        except:
            pass
        
    def update_directory(self, master, directory, currentDirVar):
        try:
            currentDirVar.set("Current Directory: " + directory)
        except:
            pass        

    def extract(self, master, selectedFolder, folderList, folderDic):
        #long-winded... but use the "pretty" formatted string to match with the folderList
        #get the un-formatted "raw" folder name string
        #use that string as a key for the folderDic
        folderListIndex = folderList[1].index(selectedFolder)
        selectedFolder = folderList[0][folderListIndex]
        selectedList =  folderDic[selectedFolder]
        print(folderListIndex)
        print(selectedFolder)
        print(selectedList)
        print(folderDic)
        #prepare csv. just clears any same name csv file.
        newCsv()
        #select mailbox(account).
        try:
            mailbox = outlook.Folders[selectedList[0]]
        except:
            pass
        #select folder.
        try:
            folder = mailbox.Folders[selectedList[1]]
        except:
            pass
        #select subfolder.
        try:
            subfolder = folder.Folders[selectedList[2]]
        except:
            pass
        #access messages object, depending on the level of the selected folder
        if len(selectedList) == 1:
            messages = mailbox.Items
        elif len(selectedList) == 2:
            messages = folder.Items
        elif len(selectedList) == 3:
            messages = subfolder.Items
        #access each message in a for loop
        for message in messages:
            for attachment in message.Attachments:
                if attachment.FileName.endswith(".txt"):
                    attachment.SaveASFile(os.getcwd() + '\\' + "temp.txt")
                    with open("temp.txt", 'r') as openFile:
                        string = openFile.read()
                        stringToCsv(string)
                    os.remove('temp.txt')

    def update_status(self, master, status, selectedFolder, folderList, folderDic):
        try:
            self.extract(master, selectedFolder, folderList, folderDic)
            status.set("Extraction complete!")
        except:
            status.set("Error. Please inform programmer. Do not use any output files.")


class MainApplication(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        self.init_window()
        
    def init_window(self):
        self.master.title("InResearch Bounced Email Extractor")
        self.buttons = Buttons(self.master)
        self.body = Body(self.master, self.buttons)

#just to create new csv each time the prog is run.
#because the stringToCsv function will only add new lines because of 'ab'
def newCsv():
    with open(outFile, 'wb') as opencsv:
        writer = csv.writer(opencsv)
        writer.writerow([datetime.date.today()])

def stringToCsv(string):
    #returns a tuple of all matches in the string
    mo = emailRegEx.findall(string)
    with open(outFile, 'ab') as opencsv:
        writer = csv.writer(opencsv)     
        for mail in list(mo):
            mail = [mail]
            writer.writerow(mail)

def getFolderList(outlook):
    #making a list within a list
    #one for formatting (pretty)
    #the other to match as a dictionary key (raw)
    menu_list_raw=[]
    menu_list_pretty=[]
    for mailbox in outlook.Folders:
        menu_list_raw.append(str(mailbox))
        menu_list_pretty.append(str(mailbox))
        for folder in mailbox.Folders:
            if len(folder.Items) > 0:
                menu_list_raw.append(str(folder))
                menu_list_pretty.append('- - ' + str(folder))
            for subfolder in folder.Folders:
                if len(subfolder.Items) > 0:
                    menu_list_raw.append(str(subfolder))
                    menu_list_pretty.append('- - - - ' + str(subfolder))
    menu_list = []
    menu_list.append(menu_list_raw)
    menu_list.append(menu_list_pretty)
    return menu_list

def getFolderDic(outlook):
    menu_dic={}
    mailbox_index = 0
    for mailbox in outlook.Folders:
        index_list = []
        index_list.append(mailbox_index)
        menu_dic[str(mailbox)] = index_list
        folder_index=0
        for folder in mailbox.Folders:
            index_list = [mailbox_index]
            index_list.append(folder_index)
            menu_dic[str(folder)] = index_list
            subfolder_index = 0
            for subfolder in folder.Folders:
                index_list = [mailbox_index, folder_index]
                index_list.append(subfolder_index)
                menu_dic[str(subfolder)] = index_list
                subfolder_index += 1
            folder_index += 1
        mailbox_index +=1
    return menu_dic



if __name__ == "__main__":
    root = tk.Tk()
#    root.geometry("400x200")
    root.resizable(0,0)
    MainApplication(root)
    root.mainloop()


        
