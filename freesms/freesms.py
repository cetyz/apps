# -*- coding: utf-8 -*-
"""
Created on Tue Aug 14 13:03:28 2018

@author: Charles Edmond TanYZ
"""

# for python 3.x

import os
import time
import csv
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

'''
driver = webdriver.Chrome('F:\WinPython\python3.6\selenium\chromedriver.exe')
driver.get('http://www.afreesms.com/freesms/')
driver.close()
'''

global driver

def message(name, token):
    message = (
'''Dear ITE Graduate,

Click to access the ITE Survey:

https://survey.inresearch.com.sg/index.php/survey/index/sid/836195/token/''' + token + '''/lang/en

InResearch Pte Ltd'''
)
    return message

# select the csv file and update the label for it(below)      
def select_csv(csvPathVar, csvNameVar, instructionsVar):
    selected_csv = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              filetypes =(('CSV File', '*.csv'),('All Files', '*.*')),
                                              title = 'Select CSV')
    if selected_csv == '':
        csvNameVar.set('No CSV Selected')
        csvPathVar.set('None')
        instructionsVar.set('Select a CSV with the headers Name, Number, and Token')
    else:
        csvNameVar.set(os.path.basename(selected_csv))
        csvPathVar.set(selected_csv)
        instructionsVar.set('Click "Initialize Browser" to start the browser')

# to update the instructions
def update_instructions(instructionsVar, string):
    instructionsVar.set(string)

def load_receipients(csv_file, counterVar, nameVar, numberVar, tokenVar, previewVar, instructionsVar):
    global driver
    counter = counterVar.get()
    driver.driver.get('http://www.afreesms.com/freesms/')
    with open(csv_file) as opencsv:
        reader = csv.reader(opencsv)
        rows = list(reader)
        if counter <= len(rows) - 1:
            nameVar.set(rows[counter][0])
            numberVar.set(rows[counter][1])
            tokenVar.set(rows[counter][2])
            previewVar.set(message(nameVar, tokenVar))
            update_instructions(instructionsVar, 'Enter the Verification Code and hit "Send"')
        else:
            nameVar.set('')
            numberVar.set('')
            tokenVar.set('')
            update_instructions(instructionsVar, 'SMSes sent to all recipients!')
            driver.driver.close()

def init_browser(instructionsVar):
    global driver
    driver = Driver()
    update_instructions(instructionsVar, 'Click Load Receipient to load in the first receipient')


def send_button(csv_file, counterVar, nameVar, numberVar, tokenVar, verificationCodeVar, instructionsVar):
    counterVar.set(counterVar.get() + 1)
    global driver
    driver.select_country()
    driver.enter_number(numberVar.get())
    driver.enter_message(nameVar.get(), tokenVar.get())
    driver.enter_verification(verificationCodeVar.get())
    driver.send()
    
    update_instructions(instructionsVar, 'Sent! Hit Load Receipients for the next receipient')

# functions related to selenium
class Driver:
    def __init__(self):
        self.driver = webdriver.Chrome('F:\WinPython\python3.6\selenium\chromedriver.exe')
    
    def select_country(self):
        countryNameBox = self.driver.find_element_by_xpath('//*[@id="smsform"]/table/tbody/tr[2]/td[2]/select')
        countrySelect = Select(countryNameBox)
        countrySelect.select_by_visible_text('Singapore')
    
    def enter_number(self, number):
        numberBox = self.driver.find_element_by_xpath('//*[@id="smsform"]/table/tbody/tr[3]/td[2]/input')
        numberBox.send_keys(str(number))
    
    def enter_message(self, name, token):
        textBox = self.driver.find_element_by_xpath('//*[@id="smsform"]/table/tbody/tr[4]/td[2]/textarea')
        textBox.send_keys(message(name, token))
    
    def enter_verification(self, number):
        verificationBox = self.driver.find_element_by_xpath('//*[@id="smsform"]/table/tbody/tr[6]/td[2]/input')
        verificationBox.send_keys(str(number))
    
    def send(self):
        sendButton = self.driver.find_element_by_xpath('//*[@id="smsform"]/table/tbody/tr[8]/td[2]/input[@id="submit"]')
        sendButton.click()


# GUI

class MainApplication(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        self.init_window()
        
    def init_window(self):
        self.master.title('SMS Sender')
        self.variables = Variables(self.master)
        self.buttons = Buttons(self.master, self.variables)
        self.body = Body(self.master, self.buttons, self.variables)
        
class Body(tk.Frame):
    def __init__(self, master, buttons, variables):
        
        csvNameLabel = tk.Label(master, textvariable = variables.csvNameVar)
        csvNameLabel.grid(row=0, column=1)
        instructionsLabel = tk.Label(master, textvariable = variables.instructionsVar)
        instructionsLabel.grid(row=1, column=0, columnspan=2)
        nameLabel = tk.Label(master, text = 'Name of recipient:')
        nameLabel.grid(row=3, column=0)
        nameVar = tk.Label(master, textvariable = variables.curNameVar)
        nameVar.grid(row=3, column=1)
        numberLabel = tk.Label(master, text = 'Number of recipient:')
        numberLabel.grid(row=4, column=0)
        numVar = tk.Label(master, textvariable = variables.curNumVar)
        numVar.grid(row=4, column=1)
        tokenLabel = tk.Label(master, text = 'Token of recipient:')
        tokenLabel.grid(row=5, column=0)
        tokenVar = tk.Label(master, textvariable = variables.curTokenVar)
        tokenVar.grid(row=5, column=1)
        previewLabel = tk.Label(master, text = 'Preview:')
        previewLabel.grid(row=6, column=0)
        previewLabel = tk.Label(master, textvariable = variables.messagePreviewVar)
        verificationLabel = tk.Label(master, text = 'Enter Verification Code:')
        verificationLabel.grid(row=8, column=0)
        
        
        
class Buttons(tk.Frame):
    def __init__(self, master, variables):
        selectCsvButton = tk.Button(master, text = 'Select CSV', command = 
                                    lambda: select_csv(variables.csvPathVar, variables.csvNameVar, variables.instructionsVar))
        selectCsvButton.grid(row=0, column=0, sticky = tk.W + tk.E)
        
        initBrowserButton = tk.Button(master, text = 'Initialize Browser', command =
                                      lambda: init_browser(variables.instructionsVar))
        initBrowserButton.grid(row=2, column=0, sticky=tk.W+tk.E)
        
        loadButton = tk.Button(master, text = 'Load Receipient', command =
                                lambda: load_receipients(variables.csvPathVar.get(), 
                                               variables.counterVar, 
                                               variables.curNameVar, 
                                               variables.curNumVar, 
                                               variables.curTokenVar,
                                               variables.messagePreviewVar,
                                               variables.instructionsVar))
        loadButton.grid(row=2, column=1, sticky = tk.W+tk.E)
        
        verificationEntry = tk.Entry(master, textvariable = variables.verificationCodeVar)
        verificationEntry.grid(row=8, column=1, sticky = tk.W+tk.E)
        sendButton = tk.Button(master, text = 'Send', command =
                               lambda: send_button(variables.csvPathVar.get(), 
                                               variables.counterVar, 
                                               variables.curNameVar, 
                                               variables.curNumVar, 
                                               variables.curTokenVar,
                                               variables.verificationCodeVar,
                                               variables.instructionsVar))
        sendButton.grid(row=9, column=1, sticky = tk.W+tk.E)
    
class Variables(tk.Frame):
    def __init__(self, master):
        self.csvPathVar = tk.StringVar()
        self.csvPathVar.set('None')
        self.csvNameVar = tk.StringVar()
        self.csvNameVar.set('No CSV Selected')
        self.instructionsVar = tk.StringVar()
        self.instructionsVar.set('Select a CSV with the headers Name, Number, and Token')
        self.curNameVar = tk.StringVar()
        self.curNumVar = tk.StringVar()
        self.curTokenVar = tk.StringVar()
        self.messagePreviewVar = tk.StringVar()
        self.verificationCodeVar = tk.StringVar()
        self.counterVar = tk.IntVar()
        self.counterVar.set(1)
        
    


if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root)
    root.mainloop()
#countrySelect = Select(countryNameBox)
