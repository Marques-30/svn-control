import webbrowser
import os
import svn
import fileinput
import sys
import subprocess
from time import strftime
from pathlib2 import Path
import getpass
import csv
import itertools
import win32com.client
from win32com.client import Dispatch, constants

user = getpass.getuser()
#choice = raw_input("Would you like this program to run automatically or manual control (Please Choose A or M):\n")

CF = raw_input("Which folder would you like to copy? ")
p = subprocess.Popen(["svn", "list", "https://subversion/repos/BIDM/"], stdout=subprocess.PIPE)
print p.communicate()
copy_envir = raw_input("Where you are directing the folder.\nEnter the name of the Environment: ")
p = subprocess.Popen(["svn", "list", "https://subversion/repos/BIDM/"+copy_envir], stdout=subprocess.PIPE)
print p.communicate()
copy_file = raw_input("Enter the name of the Target Folder: ")
message = raw_input("What is your comment for this file:\n")
file = os.listdir("C:\Users\\"+user+"\\Desktop\\"+CF)
file_data = str(file).split("'")
p = subprocess.Popen(["svn", "import", "C:\Users\\"+user+"\\Desktop\\"+CF, "--message", message, "https://subversion/repos/BIDM/"+copy_envir+"/"+copy_file+"/"], stdout=subprocess.PIPE)
header = ['Lan ID', 'File Name', "Date of Action", "Comment"]


def tried():
	try:
		path = Path("C:\\Users\\"+user+"\\Desktop\\log.csv")
		path.resolve()
	except OSError:
		Log_Document = open("log.csv", "wb")
		writer = csv.writer(Log_Document, lineterminator='\n')
		writer.writerow((header))
		writer.writerow((Data1))
		try:
			writer.writerow((Data2))
		except NameError:
			email()
		else:
			try:
				writer.writerow((Data3))
			except NameError:
				email()
			else:
				try:
					writer.writerow((Data4))
				except NameError:
					email()
				else:
					try:
						writer.writerow((Data5))
					except NameError:
						email()
					else:
						try:
							writer.writerow((Data6))
						except NameError:
							email()
						else:
							try:
								writer.writerow((Data7))
							except NameError:
								email()
							else:
								try:
									writer.writerow((Data8))
								except NameError:
									email()
								else:
									try:
										writer.writerow((Data9))
									except NameError:
										email()
									else:
										try:
											writer.writerow((Data10))
										except NameError:
											email()

	else:
		with open(r'log.csv', 'a') as f:
			writer = csv.writer(f)
			writer.writerow((Data1))
			try:
				writer.writerow((Data2))
			except NameError:
				email()
			else:
				try:
					writer.writerow((Data3))
				except NameError:
					email()
				else:
					try:
						writer.writerow((Data4))
					except NameError:
						email()
					else:
						try:
							writer.writerow((Data5))
						except NameError:
							email()
						else:
							try:
								writer.writerow((Data6))
							except NameError:
								email()
							else:
								try:
									writer.writerow((Data7))
								except NameError:
									email()
								else:
									try:
										writer.writerow((Data8))
									except NameError:
										email()
									else:
										try:
											writer.writerow((Data9))
										except NameError:
											email()
										else:
											try:
												writer.writerow((Data10))
											except NameError:
												email()


#def auto():
#	Document = raw_input("Enter the name of the CSV document: ")
#	Object = open(Document+".csv", "r")
#	log = str(Object.read()).split(",")
#	cut = log[0].split("/")
#	CF = raw_input("Which folder would you like to copy? ")
#	copy_envir = cut[0]
#	p = subprocess.Popen(["svn", "list", "https://subversion/repos/BIDM/"+copy_envir], stdout=subprocess.PIPE)
#	copy_file = cut[1]
#	message = raw_input("What is your comment for this file:\n")
#	p = subprocess.Popen(["svn", "import", "C:\Users\\"+user+"\\Desktop\\"+CF, "--message", message, "https://subversion/repos/BIDM/"+copy_envir+"/"+copy_file+"/"], stdout=subprocess.PIPE)
#	header = ['Lan ID', 'File Name', 'Revision', "Date of Action", "Comment"]
#	Data = [user, CF, 'Grab SVN Revision', strftime("%Y-%m-%d %H:%M:%S"), message]
#	try:
#		path = Path("C:\\Users\\"+user+"\\Desktop\\log.csv")
#		path.resolve()
#	except OSError:
#		Log_Document = open("log.csv", "wb")
#		writer = csv.writer(Log_Document, lineterminator='\n')
#		writer.writerow((header))
#		writer.writerow((Data))
#	else:
#		with open(r'log.csv', 'a') as f:
#		    writer = csv.writer(f)
#		    writer.writerow((Data))
#	print "file copied\n"

def email():
	print "file copied\n"
	print "\nWho would you like to send the log file to?\n"
	const = win32com.client.constants
	olMailItem = 0x0
	obj = win32com.client.Dispatch("Outlook.Application")
	newMail = obj.CreateItem(olMailItem)
	First_Name = raw_input("First Name: ")
	Last_Name = raw_input("Last Name: ")
	newMail.Subject = "SVN Mirgation"
	newMail.Body = "Attached is a log file of recent uploads to SVN from "+strftime("%Y-%m-%d %H:%M:%S")+"." 
	newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
	#newMail.HTMLBody = "<HTML><BODY>Enter the <span style='color:red'>message</span> text here.</BODY></HTML>"
	newMail.To = First_Name+"."+Last_Name+"@bankofthewest.com; "
	attachment1 = r"C:\\Users\\"+user+"\\Desktop\\log.csv"
	newMail.Attachments.Add(Source=attachment1)
	newMail.display()
	newMail.send
	print "File has been sent"
	close = raw_input("Program has ended, press enter to close ")
	sys.exit()


if __name__ == '__main__':
	Data1 = [user, file_data[1], strftime("%Y-%m-%d %H:%M:%S"), message]
	try:
		Data2 = [user, file_data[3], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data3 = [user, file_data[5], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data4 = [user, file_data[7], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data5 = [user, file_data[9], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data6 = [user, file_data[11], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data7 = [user, file_data[13], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data8 = [user, file_data[15], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data9 = [user, file_data[17], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
	try:
		Data10 = [user, file_data[19], strftime("%Y-%m-%d %H:%M:%S"), message]
	except IndexError:
		tried()
