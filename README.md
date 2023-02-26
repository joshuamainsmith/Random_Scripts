# Random_Scripts
Just some random scripts I've written for various purposes

# Jabber_Grabber.au3
Iterates through a text file of usernames to look up the user in Jabber to grab some info found there to store into a CSV file.
I made this script when I was given the task of looking up around 500 users manually.

# PrintJobs.ps1
Queries remote print servers to check if the printer spooler on each server is running. If not, then it's started.
It also checks if it appears that the print queue is stuck on the server. If so, then the spooler is restarted. If the print queue is still stuck the next time the server is reached (about 15 minutes), then the first print job is deleted and the print spooler is restarted again.
There have been some issues with the spooler stopping randomly on some of the servers and the print queue stopping.

# StudentInfoToSpreadsheet.au3
Grabs stuent info from CNS and pastes that info into a student computer tracking Excel spreadsheet. It then takes that info and pastes it into a shipping form on UPS.
Basically, this is something I wrote up to automate some of the computer packaging and shipping stuff I do.

# ps.ps1
A PowerShell script I wrote that takes a CSV file as input and writes the info I'm intersted in keeping to a text file.
Comes complete with a GUI.
