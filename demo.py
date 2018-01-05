import webbrowser
import os
import svn
import fileinput
import sys
import subprocess
from time import strftime
from pathlib2 import Path
import getpass
import time
import csv
import itertools
import win32com.client
from win32com.client import Dispatch, constants

if __name__ == '__main__':
	def email():
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
		print"\n"


	def tried():
		try:
			path = Path("C:\\Users\\"+user+"\\Desktop\\log.csv")
			path.resolve()
		except OSError:
			header = ['Lan ID', 'File Name', "Date of Action", "Comment"]
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


	target = raw_input("Enter the name for the branch folder you would like to create: ")
	message = raw_input("What is your comment for this folder:\n")
	user = getpass.getuser()
	limit = 1
	Object = open("Object-list.csv")
	reader = csv.reader(Object, delimiter=',')
	for row in reader:
		log = row[0:3]
		cut=log[0].split("/")
		count = 0
		File_source = cut[0]
		source = cut[1]
		g = subprocess.Popen(["svn", "mkdir", "https://subversion/repos/BIDM//"+File_source+"//"+source+"//"+target, "--message", message], stdout=subprocess.PIPE)
		count += 1
		g.wait()
		status = g.returncode
		if status != 0:
			Existing = raw_input("The Branch are you trying to create already exists, would you like to overwrite the branch?\nY/N: ")
			if Existing.lower() == "y":
				g = subprocess.Popen(["svn", "info", "https://subversion/repos/BIDM//"+File_source+"//"+source+"//"+target], stdout=subprocess.PIPE)
				g.wait()
				print "Branch has been updated"
			else:
				print "Adding to existing file"
		else:
			print "New Branch has been created\n"
		choice = "local"#raw_input("Are you moving the file from local repostairy or svn: ")
		if choice.lower() == "local":
			if count == int(limit):
				Clog = log[1].split("\n")
				CF = str(Clog[0])
				#Blog = log[2].split("\n")
				#TF = str(Blog[0])
				p = subprocess.Popen(["svn", "list", "https://subversion/repos/BIDM/"], stdout=subprocess.PIPE)
				p = subprocess.Popen(["svn", "list", "https://subversion/repos/BIDM/"+File_source], stdout=subprocess.PIPE)
				file = os.listdir("C:\Users\\"+user+"\\Desktop\\"+CF)
				file_data = str(file).split("'")
				p = subprocess.Popen(["svn", "import", "--message", message, "--no-ignore", "C:\Users\\"+user+"\\Desktop\\"+CF+"\\", "https://subversion/repos/BIDM/"+File_source+"/"+source+"/"+target+"/"], stdout=subprocess.PIPE)
				p.wait()
				fault = p.returncode
				if fault != 0:
					print "The file already exists in "+target+"."
					can = raw_input("Would you like to continue with the program? Y/n: ")
					if can.lower() == "n":
						email()
						print "File has been emailed\nProgram has ended, press enter to close"
						time.sleep(5)
						sys.exit()
				else:
					print "The folder has been copied"
					p = subprocess.Popen(["svn", "list", "https://subversion/repos/BIDM//"+File_source+"//"+source+"//"+target], stdout=subprocess.PIPE)
					print p.communicate()
		else:
			Clog = log[1].split("/")
			copy_envir = Clog[0]
			copy_file = Clog[1]
			p = subprocess.Popen(["svn", "copy", "https://subversion/repos/BIDM/"+copy_envir+"/"+copy_file+"/", "--message", message, "https://subversion/repos/BIDM/"+File_source+"/"+source+"/"+target+"/"], stdout=subprocess.PIPE)
			p.wait()
			print "The folder has been copied"
			p = subprocess.Popen(["svn", "checkout", "https://subversion/repos/BIDM/"+File_source+"/"+source+"/"+target+"/"], stdout=subprocess.PIPE)
			email()
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
	
print "File has been emailed\nProgram has ended, press enter to close"
time.sleep(5)
sys.exit()
