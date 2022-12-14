# MicroManageMe
_Python Project Group 1_

The purpose of this package is to create multiple Outlook appointments for course assigments based on data inputed by the user. The appointment parameters addressed in the first vesion of this package include subject title, start date, end date and body text. Additionally, reminders and importance for the appointments are determined based on the weight associated with each course on the final grade or directly by the user.
***
**Motivation**

Our motivation to develop this course comes from the frustration we face as students manually inputting assignments or exams into our Outlook calendars at the beginning of an academic semester. This process is very tedious and is unnecessarily time consuming, so we decided to make it easier. MicroManageMe was produced using resources provided by the University of Florida for EML4930 Special Topics: Python for Engineers.
***
**Build Status**

We are unaware of any bugs and would welcome any testing by the GitHub community. We have only tested this package on 3 machines with 2 Outlook accounts, so one concern is the automatic use of Outlook with the win32 API. So far in our testing, we have not ran into an issue with Outlook credentials. If appointments are not created automatically in your Outlook application, then please comment that bug on our repository page.
***
**Code Style**

No particular code style was implemented except for the use of camel case as a naming convention for functions.
***
**Screenshots**

Below is what you should see if the appointment creator works properly. This is an example appointment with arbitrary fields. In most cases, appointments should pop-up instantly in your Outlook page.
![Appointment Display](https://github.com/zoeestberg/MicroManageMe/blob/main/Appointment%20Display.png)
***
**Installation**

1. Microsoft Outlook is required for this package to function. An account must be signed in as well.

2. Install and extract the zip file containing the github repository into a file location of your choice.

3. Run the powershell application in the same security context as your outlook (administrator, base privilege, etc).

4. Set your location in powershell to the folder containing the contents of the github repository.
`Set-Location 'Folder Directory'`

5. Import the yml file as the working environment in powershell, activate the environment.
`conda env create -n MicroManageMe --file .\MicroManageMe.yml`
`conda activate MicroManageMe`

6. Run the MicroManageMe.py script with the arguments `-c 'config filename'` and `-s 'course schedule filename'`.
`python .\MicroManageMe.py -c 'Param Config.csv' -s 'Course 1.csv'`

This is an example screenshot of what should be input into your Powershell to functionally use the script.
![Powershell Example Setup](https://github.com/zoeestberg/MicroManageMe/blob/main/Powershell%20Example%20Setup.PNG)

***
**How to use?**

Refer to docstrings for detailed descriptions of individual functions. The input format for the configuration files is as follows:

*Param Config.csv*

- `Low Priority Boundary` Whole number identifying the % weight of an assignment in a course designating a low priority assignment. (i.e. an assignment is 5% of grade but low priority boundary is 10%, assignment is assigned internally as low priority).
- `Medium Priority Boundary` Whole number identifying the % weight of an assignment in a course designating a low priority assignment. Note, high importance is designtaed automatically as any value above the medium importance boundary.
- `Low Importance Reminder` Whole number of days before appointment of low importance to be reminded of event.
- `Medium Importance Reminder` Whole number of days before appointment of medium importance to be reminded of event.
- `High Importance Reminder` Whole number of days before appointment of high importance to be reminded of event.
- `Time Zone` 3 capital letters identifying time zone (i.e. Eastern Standard Time -> EST).

*Course 1.csv*

Please note, each row following the headers corresponds to a single assignment. An unlimited number of rows can be added to the configuration file, but there should not be any gaps between rows.

- `Assignment Title` Title of the assignment you are assigning an Outlook event to. This parameter will function as the subject of the appointment.
- `Assignment Type` There are 3 assignment types "HW", "Assignment" or "Exam". These are just classifiers, or an addded descriptor for an event.
- `Weight` Percentage of final course grade the assignment is worth in a whole number (integer). The weight in combination with the prioriy boundaries automatically assign appointment reminders.
- `Due Date` Date the assignment is due in MM/DD/YYYY format. Note, only dates in the future can be assigned to an appointment (i.e. December 25, 2050 -> 12/25/2050).
- `URL or Description`: This parameter is added to the appointment body text. Any information relevant to the assignment can be included in this section, including linkable URLs.

***
**Credits**
|Contributor|Email|
|----------------|-----|
|Zoe Estberg|zoeestberg@ufl.edu|
|Elijah Rice|elijahrice@ufl.edu|
|Samuel Roshaven|sroshaven@ufl.edu|

This package was developed under the supervision of Dr. Salil Bavdekar at the University of Florida.
