###############################################################################################################################################
## This script copies all of the Calendar entries from the Lotus Notes Calendar to you Outlook Calendar                                      ##
## Author: Thomas Land																														 ##
## E-Mail: tland@outlook.com                                                                                                                 ##
## Twitter: @thomaslandjr        																											 ##
## I am not responsible for code that you have changed yourself, if you do not understand what a change will do, leave it alone!             ##
## Refer to the documentation sent with this code to configure it yourself.																	 ##
###############################################################################################################################################


# The following code will import the necessary libraries to run this script.
from win32com.client import makepy
from win32com.client import Dispatch
from win32com.client import constants
import getpass, os, tempfile, re

# The following code will generate the objects that this script needs to interact with Lotus Notes
makepy.GenerateFromTypeLibSpec('Lotus Domino Objects')
makepy.GenerateFromTypeLibSpec('Lotus Notes Automation Classes')

# The following function will create a generator for Lotus Notes Calendar Entries
def iterCal(entries):
	entry = entries.GetFirstDocument()
	while entry:
		yield entry
		entry = entries.GetNextDocument(entry)
		
# The Following function will add Lotus Notes Calendar Entries to your Outlook Calendar		
def addToOutlook(subject, location, body, startTime, endTime, required):	
		cal = Dispatch('Outlook.Application')
		namespace = cal.GetNamespace("MAPI")
		outApt = namespace.GetDefaultFolder(9).Items
		for apts in outApt:
			if subject in apts.Subject and str(apts.Start) == startTime:
				print "Meeting already exists in Outlook"
				return
		apt = cal.CreateItem(1) 
		apt.Start = startTime
		apt.End = endTime
		apt.RequiredAttendees = required
		apt.Subject = subject
		apt.Location = location
		apt.Body = body
		apt.ReminderSet = True
		apt.ReminderMinutesBeforeStart = 15
		apt.Save()

#The following code will setup the initial information that Lotus Notes needs to Connect to the Notes Server
notes = Dispatch('Lotus.NotesSession')
mailServer = raw_input("What is your mail server: ")
mailPath = raw_input("What is your mail path: ")
mailPassword = getpass.getpass()
try:
	notes.Initialize(mailPassword)
	notesDB = notes.GetDatabase(mailServer, mailPath)
except:
	raise Exception("Cannon access mail using %s on %s" % (mailPath, mailServer))
	
# The following line gets the Calendar from Lotus Notes
cal = notesDB.GetView("$Calendar")

# The following line sets up an empty list to hold recurring Calendar appointments	
added = []

# The rest of the code below does the brunt of the work fetching individual Lotus Notes Calendar entries and sending them to the addToOutlook function
for entry in iterCal(cal):
	subject = ''.join(entry.GetItemValue("Subject"))
	location = ''.join(entry.GetItemValue("Location"))
	required = ''.join(entry.GetItemValue("RequiredAttendees"))
	body = entry.GetItemValue("Body")[0]
	body = body + '\r\n\r\nThis was imported from your Lotus Notes Calendar, you will need to check there for any attachments that may exist!' 
	doesRepeats = entry.GetItemValue("Repeats")
	startTime = entry.GetItemValue("StartDateTime")[0]
	startTime = str(startTime)
	endTime = entry.GetItemValue("EndDateTime")[0]
	endTime = str(endTime)
	if doesRepeats[0] == '1':
		aptObject = {subject:startTime}
		if aptObject not in added:
			for appointment in entry.GetItemValue("StartDateTime"):
				startTime = str(appointment)
				endTime = str(endTime).replace(str(endTime)[0:8], str(appointment)[0:8])
				addToOutlook(subject, location, body, startTime, endTime, required)
			added.append(aptObject)
		else:
			print "Already added this Recurring Appointment to Calendar"
	else:
		addToOutlook(subject, location, body, startTime, endTime, required)
		
print "Successfully Moved Lotus Notes Calendar Entries to your Outlook Calendar!"
	
