# Call Quality Dashboard Report

# Description

Script will provide total stream count including audio, video, app sharing for provided start time and end time, CQD report of a given time

PowerShell should be more than 3.0 version

# Example

##### Example 1 for total stream count including audio, video, app sharing

![Example](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CQD-Example.png)

##### Example 2 for CQD report of the given time (DD-MM-YYYY(Ex:31-03-2020)

Start Date: 1-10-2020

End Date: 1-11-2020

 # Parameters
 
 `-Date`
 
 Type: String 
 
 # Inputs
 
  Provide input 1 to get total stream count including audio, video, app sharing
  
   Start Date – “Please provide start date"
   
   End Date – “Please provide end date"
  
  Provide input 2 to get the CQD report of a given time
  
  Give the input file as shown below. Keep this file in current location(CQD_Input.csv). The script will collect the data from `CQD_Input.csv` file and capture the data from call     quality dashboard

 |Dimensions  |	Measures| OutPutFilePath |	StartDate| EndDate | OutPutType	| MediaType	| IsServerPair |
 |------------|---------|----------------|-----------|---------|------------|-----------|--------------|

 Each row data will be collected and executed through script accordingly
  
 To construct the input.csv file refer [dimensions-and-measures-available-in-call-quality-dashboard](https://docs.microsoft.com/en-us/microsoftteams/dimensions-and-measures-available-in-call-quality-dashboard)
 
 # Procedure
 
PowerShell should be more than 3.0 version

Run the script

Once you run the script it will prompt for option 1 or 2

If you have chosen the **option 1** please provide the parameters 

Start Date – “Please provide start date” 

End Date – “Please provide end date” 

Press enter to continue 

Or if you have chosen the **option 2** please provide the `Input.csv` file 

Now the script will pop-up for Teams Administrator credentials to connect the CQD tool

![Signin](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CQD-Signin.png)

Provide the Teams Administrator credentials

# Output

For option 1 

##### Example output

Script will execute and creates `cqdoutput.csv` file
![SampleOutput](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CQD-SampleOutput.png)

For option 2

##### Example output

![Output](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/CQD-output.png)

A log file will be generated with exceptions, errors along with script execution time
