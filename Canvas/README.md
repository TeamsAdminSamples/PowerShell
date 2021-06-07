# Canvas

# Description

Downloading the synchronization files from data source Canvas and uploading the data to the school data synchronization tool

# Index

**System requirements**

**Preparing the system**   

**Getting the Token ID**

# Downloading Data from Canvas

	Without sync.csv file

	With sync.csv file

# File Upload

Canvas_FileUpload (Manual Script)

Canvas_FileUpload (Automate)

# System requirements

A suitable version of Windows PowerShell is available for these operating systems:

* Windows 10

* Windows 8.1 Pro

* Windows 8.1 Enterprise

* Windows 7 SP1

* Windows Server 2019

* Windows Server 2016

* Windows Server 2012 R2

* Windows Server 2008 R2 SP1

# Preparing the system

Go to start button right click on the windows PowerShell Run as Administrator

Check the ExecutionPolicy by running below command in the console

  	Get-ExecutionPolicy 

Set the ExecutionPolicy to Unrestricted

	Set-ExecutionPolicy -ExecutionPolicy Unrestricted 

# Getting the Token ID

Generate the new access token using the below process

login into the Canvas website [uri](Ex:https://jlg.instructure.com)

Go to Account &rightarrow; settings &rightarrow; Approved integrations &rightarrow; New Access Token

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasApproved%20Integrations.png)

# Downloading Data from Canvas

**Without sync.csv file**

1. How to run the script?

  	1. Run the script by double-clicking “CanvasApiMainScript.ps1”

  	2. Script will ask for Uri - Pass login canvas [Uri](Ex: https://jlg.instructure.com)

  	3. Script will ask fir Token ID - Provide created Token id (Ex:4626~OhR7Arue1IhsaJnplrLYYeqj********)

2. Where files will be created?

  	Files will be created in the same directory where the script file is located
    
Once you run the script, it downloads the all available courses data from canvas and creates courses.csv and looks for sync.csv file ( If location not having sync.csv file 
Script will create sync.csv and ask you for the provide the details to move further)

Script will ask the user to “keep course id in sync.csv file and press Y to proceed”

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasContinue.png)

Hit “Y” and enter. The remaining script will run and create the below files for given input `sync.csv` course ids

3. What files will be created?

	School.csv

	Section.csv

	Teacher.csv

	Student.csv

	StudentEnrollment.csv

	TeacherRoster.csv

**With sync.csv file**

1. How to run the script?

 * Run the script by double-clicking `CanvasApiMainScript.ps1`

 * Script will ask for Uri - Pass login [canvas Uri](https://jlg.instructure.com)

 * Script will ask fir Token ID - Provide created Token id (Ex:4626~OhR7Arue1IhsaJnplrLYYeqj********)

2. Where files will be created?

	Files will be created in the same directory where the script file is located

Once you run the script, it downloads the all available courses data from canvas and creates courses.csv

Script will check sync.csv file. If the file is available runs the remaining script. If the file is not available script asks you to re-run the script by keeping the sync.csv file in the current location

Script will create below files taking the input from sync.csv

	School.csv

	Section.csv

	Teacher.csv

	Student.csv

	StudentEnrollment.csv

	TeacherRoster.csv

# Canvas Files Upload to SDS:

**Canvas_FileUpload (Manual Script):**

Uploading files to School Data Sync (Manual process):

* [Link](https://sds.microsoft.com) 

* Login using your Global Admin account

* Click on Add Profile

* Enter sync profile Name

* Choose Sync method &rightarrow; Upload CSV files

* Choose Type for Csv files you are using &rightarrow; CSV files: SDS Format

* Click start

* Sync options: choose new users or Existing users

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasSyncOptions.png)

* Upload files: Select created 6 csv files (School.csv, Section.csv, Teacher.csv, Student.csv, StudentEnrollment.csv, TeacherRoster.csv)

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/Canvasdatafile.png)

* Select “Replace unsupported special characters” option 

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasReplaceUnsupported.png)

* When should we stop syncing this profile?

Select the date 

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasSelect%20the%20date.png)

* Choose Teacher Options

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasChooseTeacher.png)

Select licenses for Teachers

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasChooseTeacherLicense.png)

* Choose Student Options

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasStudentption.png)

Select licenses for Students

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasStudentLicense.png)

* Review all data and select Create a profile

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasPleasewait.png)

* Once Sync profile is submitted it will take some time to Create sync profile

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasSettingUp.png)

Once it is created. It will show the status of sync

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasValidatingFiles.png)

It will create new Teams in Teams Console

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasNewdataabove.png)

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasNew.png)

Once it's done. It will create O365 groups as shown in below same will reflect in Teams

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasNew2.png)

# Canvas File Upload to SDS (Automation):

Keep the below files in the current folder where you are running the script

azcopy.exe

uploading .csv files (School.csv; section.csv; student.csv; studentEnrollment.csv; teacher.csv; teacherroster.csv)

Navigate to current folder Run the script using run as admin rights

ex: if I want to run the script in d:

![](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CanvasTochangethepath.png)

Script will run in the current folder and create (conf. JSON) file and sastoken.cmd files

Provide the username, password and syncprofileName for the first time running the script

Second time onwards script will provide sync status

# FAQs/Common Problems:

Active Directory sync error

If teachers or students are invited but they don’t enrol as teacher or student. They will not be available as a teacher or student in the created team

If No Teachers – SDS won’t create teams for that section (if it has the students also)

In tenant, Teacher will be assigned teacher license and the student will be student license
