#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import numpy as np
import datetime
import win32com.client as client
import pytz
import argparse


# In[3]:


parser = argparse.ArgumentParser()
parser.add_argument('-c', '--config', action='store', dest='configName', help='The name of the config file')
parser.add_argument('-s', '--schedule', action='store', dest='scheduleName', help='The name of the user schedule file')
parser.add_argument("-f", "--fff", help="a dummy argument to fool ipython", default="1")

args = parser.parse_args()
configName = args.configName
scheduleName = args.scheduleName


# In[4]:


importanceTime = np.array([0, 0]) #Initialize array
  
def importanceLevel(): #Establishes 2x1 array with max boundary conditions for low and medium priority tasks
  importanceTime[0] = int(input("Enter maximum work-hours for Low-Priority task: "))
  importanceTime[1] = int(input("Enter maximum work-hours for Medium-Priority task: "))
  
  return


# In[5]:


importanceLead = {'l' : 2, 'm' : 5, 'h' : 14} #Initialize dictionary
    
def reminderLead():
    importanceLead['l'] = int(input('Enter reminder lead time in days for low-priority tasks: '))
    importanceLead['m'] = int(input('Enter reminder lead time in days for medium-priority tasks: '))
    importanceLead['h'] = int(input('Enter reminder lead time in days for high-priority tasks: '))
    
    return 


# In[6]:


def dateConverter(dateString): #Returns date as a datetime object
    #split string into date and month
    date = dateString.split("/")
    
    #Create Datetime object using date
    dateObject = datetime.datetime(int(date[2]), int(date[0]), int(date[1]), hour=9)
    dateObject = pytz.timezone('US/Alaska').localize(dateObject)
    #Add the year to the beginning of the array
    
    return dateObject

def priorityEvaluator(weighting, weightRanges): #Returns a string 'h', 'm', 'l', based on what range the expected hours falls into

    #If else change to determine priority
    if weighting < weightRanges[0]:
        priority = 'l'
    elif weighting >= weightRanges[0] and weighting < weightRanges[1]:
        priority = 'm'
    elif weighting >= weightRanges[1]:
        priority = 'h'
    
    return priority

def createDeltaTime(leadTimes, priority): #Input the array of leadtimes and the priority to return a deltatime object

    #Determine amount of leadtime based on the priority
    if priority == 'l':
        days = leadTimes['l']
    elif priority == 'm':
        days = leadTimes['m']
    elif priority == 'h':
        days = leadTimes['h']
    
    #Return the timedelta object
    return datetime.timedelta(days = int(days))

def createStartTime(dateObject, deltaTime): #Input the due date datetime object and the deltatime object to create a start datetime object
    return dateObject - deltaTime
    
def createDF(userDF, weightingRanges, leadTimes): #Input csv df and other parameters to create df to be passed to o365 module
    #initialize df with columns
    finalDF = pd.DataFrame(columns =['Assignment Title', 'Assignment Type', 'Due Date', 'Start Date', 'URL'])
    
    #Iterates over original dataframe
    for row in range(len(userDF)):
        #Creates due date datetime
        dueDate = dateConverter(userDF.loc[row, 'Due Date'])
        
        #Creates priority
        priority = priorityEvaluator(userDF.loc[row, 'Weight'], weightRanges)
        
        #Creates start time
        startDate = createStartTime(dueDate, createDeltaTime(leadTimes, priority))
        
        #Creates new row for o365 Data Table
        newRowColumns = ['Assignment Title', 'Assignment Type', 'Due Date', 'Start Date', 'URL']
        newFrame = pd.DataFrame(columns=newRowColumns)
        newRowValues = [userDF.loc[row, 'Assignment Title'], userDF.loc[row, 'Assignment Type'], dueDate, startDate, userDF.loc[row, 'URL or Description']]
        
        #Creates new DF for row of values
        newFrame.loc[0] = newRowValues

        #Appends row
        finalDF = pd.concat([finalDF, newFrame], ignore_index = True)
        
    return finalDF


# In[7]:


userDF = pd.read_csv(scheduleName)
userConfig = pd.read_csv(configName)

weightRanges = [userConfig.loc[0, 'Low Priority Boundary'], userConfig.loc[0, 'Medium Priority Boundary']]
importanceLead = {'l' : userConfig.loc[0, 'Low Importance Reminder'], 'm' : userConfig.loc[0, 'Medium Importance Reminder'], 'h' : userConfig.loc[0, 'High Importance Reminder']}
finalDF = createDF(userDF, weightRanges, importanceLead)


# In[8]:


outlook = client.Dispatch("outlook.application")

def event(subject,body,end,start):
    appt = outlook.CreateItem(1) # AppointmentItem
    appt.Start = start # yyyy-MM-dd hh:mm
    appt.Subject = subject
    appt.body = body
    appt.end = end
    appt.Save()

for row in range(len(finalDF)):
    sub = finalDF.loc[row, 'Assignment Title']
    bod = finalDF.loc[row, 'URL']
    en = finalDF.loc[row, 'Due Date'].strftime("%m/%d/%Y %H:%M:%S") + " AM"
    st = finalDF.loc[row, 'Start Date'].strftime("%m/%d/%Y %H:%M:%S") + " AM"
    event(sub, bod, en, st)
    


# In[ ]:




