# SVN Automation script

This python script automates svn as well as local computer files by copying and moving them between the two. It also displays the differences between the two branches and folders.

### Installation

For this code to work you will need to install python 2.7, shutil, svn, csv, subprocess, fileinput, and svn tortoise. To install python package simply enter the text below into command line. And to install svn tortoise follow the link to downlaod the latest version: https://tortoisesvn.net/downloads.html

	pip install webbrowser
	pip install os
	pip install svn
	pip install fileinput
	pip install sys
	pip install subprocess
	pip install getpass
	pip install pypiwin32

### Start

First make sure the files are placed on the desktop, to start the program double click the file. It will ask for what you want to name the Branch folder, enter the name then press enter. Next it should ask what you want to place as a comment, enter the comment you wish to place then press enter. It will then upload a local folder labeled in object-list and import it within SVN. Finally it will ask for the First and Last name of the person you want to send a report of the upload to.


### Results

The new should be visible within https://subversion/svnwebclient/directoryContent.jsp?location=BIDM within the Comment section your see the new text you have linked to the files as well as it's age equaling 0. As well as a Report created on the desktop stating which files were uploaded, by who, and the time that was also emailed to the recipient. 


### Contributors 

Jessie Zavala
Amit Gohel