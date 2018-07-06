# -*- coding: utf-8 -*-
"""
Created on Mon Apr 16 17:37:11 2018

@author: Charles Edmond TanYZ

Simple app to:
1. Log in to Limesurvey and get your key to perform operations
2. Find a select a csv to change emails
3. Change the emails
"""
#####
# for the csv, first column will be the "key", second column will be the "value"
#####

import Tkinter as tk
import urllib
import urllib2
import json
import sys
import csv
import os
import tkFileDialog



###############################################################################
###############################################################################

def get_session_key(user, pw):
    req = urllib2.Request(url='https://survey.inresearch.com.sg/index.php/admin/remotecontrol',\
                          data='{\"method\":\"get_session_key\",\"params\":[\"'+user+'\",\"'+pw+'\"],\"id\":1}')
    req.add_header('content-type', 'application/json')
    req.add_header('connection', 'Keep-Alive')
    try:
        f = urllib2.urlopen(req)
        myretun = f.read()
        #print myretun
        j=json.loads(myretun)
        return j['result']
    except :
        e = sys.exc_info()[0]
        print ( "<p>Error: %s</p>" % e )
	
def release_session_key(relkey):
    req = urllib2.Request(url='https://survey.inresearch.com.sg/index.php/admin/remotecontrol',\
                          data='{\"method\":\"release_session_key\",\"params\":[\"'+relkey+'\"],\"id\":1}')
    req.add_header('content-type', 'application/json')
    req.add_header('connection', 'Keep-Alive')
    try:
        f = urllib2.urlopen(req)
        myretun = f.read()
        #print myretun
        j=json.loads(myretun)
        return j['result']
    except :
        e = sys.exc_info()[0]
        print ( "<p>Error: %s</p>" % e )

def get_participant_email(skey, sid, tid):
    req = urllib2.Request(url='https://survey.inresearch.com.sg/index.php/admin/remotecontrol',\
                      data='{\"method\":\"get_participant_properties\",\
                              "params\":\
                              [\
                              "'+skey+'\",\
                              "'+sid+'\",\
                              "'+tid+'\",[\
                              "email\"]\
                              ],\
                              "id\": 1}')
    req.add_header('content-type', 'application/json')
    req.add_header('connection', 'Keep-Alive')
    try:
        f = urllib2.urlopen(req)
        myretun = f.read()
        #print myretun
        j=json.loads(myretun)
        return j['result']
    except:
        e = sys.exc_info()[0]
        print ( "<p>Error: %s</p>" % e )

def set_participant_email(skey, sid, tid, newEmail):
    req = urllib2.Request(url='https://survey.inresearch.com.sg/index.php/admin/remotecontrol',\
                      data='{\"method\":\"set_participant_properties\",\
                              "params\":[\
                                         "'+skey+'\",\
                                         "'+sid+'\",\
                                         "'+tid+'\",\
                                         {\
                                         "email\":\"'+newEmail+'\",\
                                         "sent\":\"N\",\
                                         "remindersent\":\"N\",\
                                         "remindercount\":\"0\"\
                                         }\
                                         ],\
                              "id\": 1}')
    req.add_header('content-type', 'application/json')
    req.add_header('connection', 'Keep-Alive')
    try:
        f = urllib2.urlopen(req)
        myretun = f.read()
        #print myretun
        j=json.loads(myretun)
        return j['result']
    except:
        e = sys.exc_info()[0]
        print ( "<p>Error: %s</p>" % e )

def list_surveys(skey, user):
    req = urllib2.Request(url='https://survey.inresearch.com.sg/index.php/admin/remotecontrol',\
                          data='{\"method\":\"list_surveys\",\"params\":[\"'+skey+'\",\"'+user+'\"],\"id\":1}')
    req.add_header('content-type', 'application/json')
    req.add_header('connection', 'Keep-Alive')
    try:
        f = urllib2.urlopen(req)
        myretun = f.read()
        #print myretun
        j=json.loads(myretun)
        return j['result']
    except :
        e = sys.exc_info()[0]
        print ( "<p>Error: %s</p>" % e )

#hard coded to list only up to 18000 participants
def list_participants(skey, sid):
    req = urllib2.Request(url='https://survey.inresearch.com.sg/index.php/admin/remotecontrol',\
                          data='{\"method\":\"list_participants\",\
                                  "params\":[\"'+skey+'\",\"'+sid+'\",\"0\",\"10000\"],\"id\":1}')
    req.add_header('content-type', 'application/json')
    req.add_header('connection', 'Keep-Alive')
    try:
        f = urllib2.urlopen(req)
        myretun = f.read()
        #print myretun
        j=json.loads(myretun)
        return j['result']
    except :
        e = sys.exc_info()[0]
        print ( "<p>Error: %s</p>" % e )

###############################################################################
###############################################################################

# select working directory and update the label for it (below)
def select_directory():
    directory = tkFileDialog.askdirectory(title='Select folder datafile is located in...')
    try:
        os.chdir(directory)
        return(directory)
    except:
        print('Directory was not selected')
        pass
    
def update_directory(stringVar):
    directory = select_directory()
    try:
        stringVar.set('Selected Directory: ' + directory)
    except:
        print('Could not update directory')
        pass

# select the csv file and update the label for it(below)    
def select_datafile():
    datafile = os.path.basename(tkFileDialog.askopenfilename(title='Select CSV',initialdir=os.getcwd()))
    if datafile == '':
       datafile = "No file selected"
    return datafile

def update_datafile(dataVar, stringVar):
    datafile = select_datafile()
    try:
        stringVar.set('Selected CSV: ' + datafile)
        dataVar.set(datafile)
    except:
        print('Could not update to selected CSV')
        pass    

#to refresh survey list upon log in (see below)
def refresh_survey_list(menuButton, menuVar, skey, user):
    menuVar.set('')
    menuButton['menu'].delete(0, 'end')
    surveyDic = get_survey_dic(skey, user)
    new_choices = list(surveyDic)
    for choice in new_choices:
        menuButton['menu'].add_command(label=choice, command=tk._setit(menuVar, surveyDic[choice]))

# 'log in'
def update_session_key(user, pw, keyString, resultVar, menuButton, menuVar):
    key = get_session_key(user, pw)
    if type(key) == dict:
        resultVar.set(key['status'])
    else:
        try:
            keyString.set(key)
            resultVar.set('Log in success!')
            refresh_survey_list(menuButton, menuVar, key, user)
            menuVar.set('No Survey Selected')
            print(key)
        except:
            resultVar.set('Log in failure')
            print('Could not set key to variable')
            pass
        
def getDic(inFile, statusVar):
    dic ={}
    if inFile.endswith('.csv'):
        with open(inFile, 'r') as readcsv:
            reader = csv.reader(readcsv)
            for row in reader:
                try:
                    dic[row[0]] = row[1]
                except IndexError:
                    dic[row[0]] = ''
        return dic
    else:
        return None


def change_emails(master, skey, sid, statusVar, inFile):
    dic = getDic(inFile, statusVar)
    if dic == None:
        statusVar.set('You did not select a csv')
    else:
        timerCounter = 1.00
        tokenList = list_participants(skey,sid)
        for token in tokenList:
            try:
                tokenEmail = token['participant_info']['email']
            except:
                continue
            if tokenEmail in dic:
                set_participant_email(skey, sid, token['tid'], dic[tokenEmail])
                print('Token ID: ' + token['tid'])
                print(str(tokenEmail) + ' changed to: ' + dic[tokenEmail])
            currentTimer = int(timerCounter/float(len(tokenList))*100)
            if currentTimer%5== 0:
                statusVar.set(str(currentTimer)+'%')
                master.update_idletasks()
            timerCounter += 1.00
            
def get_survey_dic(skey, user):
    try:
        surveyList = list_surveys(skey, user)
        surveyDic = {}
        for survey in surveyList:
            surveyDic[survey['surveyls_title']] = survey['sid']
        return(surveyDic)
    except:
        print('could not get survey list from LS')
        return['No survey selected']
###############################################################################
###############################################################################

            
'''
changeEmails(dic)
print(release_session_key(mykey))'''


class MainApplication(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.master = master
        self.init_window()
        
    def init_window(self):
        self.master.title('Email changer')
        self.variables = Variables(self.master)
        self.buttons = Buttons(self.master, self.variables)
        self.body = Body(self.master, self.buttons, self.variables)
        
class Body(tk.Frame):
    def __init__(self, master, buttons, variables):
        
        ##### LABELS #####
        
        userLabel = tk.Label(master, text='Username: ')
        userLabel.grid(row=0, column=0)
        passLabel = tk.Label(master, text='Password: ')
        passLabel.grid(row=1, column=0)
        logInStatusLabel = tk.Label(master, textvariable=variables.logInStatusVar)
        logInStatusLabel.grid(row=1, column=2)
        dirLabel = tk.Label(master, textvariable = variables.dirVar)
        dirLabel.grid(row=2, column=1)
        csvLabel = tk.Label(master, textvariable = variables.csvTextVar)
        csvLabel.grid(row=3, column=1)
        statusLabel = tk.Label(master, textvariable= variables.statusVar)
        statusLabel.grid(row=4, column=1)

class Buttons(tk.Frame):
    def __init__(self, master, variables):
        userEntry = tk.Entry(master, textvariable=variables.userVar)
        userEntry.grid(row=0, column=1)
        passEntry = tk.Entry(master, textvariable=variables.passVar, show='*')
        passEntry.grid(row=1, column=1)

        changeDirButton = tk.Button(master, text='Select Directory', command=
                                    lambda: update_directory(variables.dirVar))
        changeDirButton.grid(row=2, column=0, sticky=tk.W+tk.E)
        selectCSVButton = tk.Button(master, text='Select CSV', command=
                                    lambda: update_datafile(variables.csvVar,variables.csvTextVar))
        selectCSVButton.grid(row=3, column=0, sticky=tk.W+tk.E)
        selectSurveyMenu = tk.OptionMenu(master, variables.surveyListVar, *['test','test'])
        selectSurveyMenu.grid(row=4,column=0)
        getKeyButton = tk.Button(master, text='Log In', command=
                                 lambda: update_session_key(userEntry.get(),passEntry.get(),variables.keyVar, variables.logInStatusVar, selectSurveyMenu,variables.surveyListVar))
        getKeyButton.grid(row=0,column=2,sticky=tk.W+tk.E)
        
        #################### just testing ########################
        changeEmailsButton = tk.Button(master, text='Change Emails', command=
                                       lambda: change_emails(master,variables.keyVar.get(),variables.surveyListVar.get(),variables.statusVar, variables.csvVar.get()))
        changeEmailsButton.grid(row=4,column=2, sticky=tk.W+tk.E)


class Variables(tk.Frame):
    def __init__(self, master):
        self.userVar = tk.StringVar()
        self.passVar = tk.StringVar()
        self.logInStatusVar = tk.StringVar()
        self.keyVar = tk.StringVar()
        self.dirVar = tk.StringVar()
        self.dirVar.set(os.getcwd())
        self.csvVar = tk.StringVar()
        self.csvTextVar = tk.StringVar()
        self.csvTextVar.set('No CSV selected')
        self.surveyListVar = tk.Variable()
        self.surveyListVar.set('Select Survey')
        self.statusVar = tk.StringVar()
        
    
if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root)
    root.mainloop()
