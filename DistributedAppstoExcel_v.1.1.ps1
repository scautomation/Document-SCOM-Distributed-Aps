
<# 
.SYNOPSIS 
 The purpose of this script is to generate an Excel spreadsheet with Distrubuted Applications as defined in System Center Operations Manager. The script will create a new work sheet
 for each Distributed Application and then get each component within the DA, then get the objects contained in each component.

 This script must be run from a computer that has Excel installed. All SCOM commands are run remotely.

.DESCRIPTION 
 Powershell script to document SCOM Distributed Applications
 
.NOTES 
Configuring Distributed Applications for clients on pretty much every SCOM deployment, also required that I document the Distributed Applicatons. Initially this script generated CSVs
and I would have to manually make a "pretty" Excel file out of them to deliver to the client. Therefore I wanted to challenge myself to try and write a complete script that does it all
for me.

This information is also useful when creating Visio based dashboards for Squaredup. To create such dashboard you need to provide the Visio Object with a SCOM Object ID so that it can 
get its health after importing the Dashboard into SquaredUp. That is why this script will also grab SCOM Object IDs and add them into the spreadsheet.

.CREDITS
Thanks to KornKolio for the Excel methods I borrowed from https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Script-Get-beced710
Thanks to @MikeLovasco for helping test
Thanks to anyone who has taken the time to blog examples, the community is amazing.
 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 02.20.2018                                                                  |
│   AUTHOR      : Billy York                                                                  |
|   TWITTER     : @scautomation                                                               |
|   BLOG        : www.systemcenterautomation.com                                              |
│   DESCRIPTION : Initial Release                                                             |
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
 
.PARAMETER SCOM - Required
This is your SCOM Management Server 
 
.PARAMETER Credential - Optional but will prompt for credentials
Credentials that have access to both the SCOM environment and powershell remoting access to the SCOM Server. 

.PARAMETER SavenAndClose - Boolean - optional
Accepts $true  

.PARAMETER Path
Path with which you want to save the document. If path does not exist script will attemp to create it.

.PARAMETER Filename 
File name to give to your Excel document. 
 

 EXAMPLE 1
.\DistributedAppsToExcel.ps1 -scom "om01.sandlot.dom" -crednetial $cred

 EXAMPLE 2
 .\DistributedAppsToExcel.ps1 -scom $scom -credential $cred

 EXAMPLE 3
 .\DistributedAppsToExcel.ps1 -scom $scom -credential sandlot\billy

 EXAMPLE 4
.\DistributedAppsToExcel.ps1 -scom $scom -credential $cred -saveandclose $true -path "c:\temp\" -filename "scomDistributedApplications" 

This example will generate the report, save it to c:\temp with a filename of scomDistributedApplications.xlsx and close the Excel document. All other scenarios will leave the document open.
You will need to manually save the report.
#> 


	param (
		# Mandatory SCOM Management Server
    	[parameter(Mandatory=$true)]
        [string]$scom,

		# Mandatory credentials for access SCOM Remotely
        [parameter(Mandatory=$true)]
        [securestring]$credential, 

        # Optional use this if you want to save the document produced
        [parameter(Mandatory=$false)]
        [boolean]$saveandclose,

        # Optional, but required if using the SaveandClose option
        [parameter(Mandatory=$false)]
        [string]$path,

        # Optional, but required if using the SaveandClose option
        [parameter(Mandatory=$false)]
        [string]$filename

    )


#Import the Operations Manager Powershell Module
Import-Module OperationsManager

#Create initial spreadsheet 
#get SCOM Management Group Name
$MG = get-scommanagementgroup -Computername $scom -credential $credential
$mgname = $mg.name

#use to remove Slashes from DA Names
$pattern = '[\\/]'

#to use with sheet items
$count = 1

#Create a new Excel object using COM 
#$sheetcount = 1 
$doc = New-Object -ComObject Excel.Application
$doc.visible = $True 
$doc.DisplayAlerts = $false
$doc.WindowState = "xlMaximized"
$Excel = $doc.Workbooks.Add() 
$MainSheet = $Excel.Worksheets.Item($count)  
$MainSheet.Name = 'Distributed Applications' 


#Create a Title for the first worksheet 
$row = 1 
$Column = 1 
$MainSheet.Cells.Item($row,$column)= "Distributed Applications for $MGname Management Group"

$range = $MainSheet.Range("a1","s2") 
$range.Merge() | Out-Null 
$range.VerticalAlignment = -4160 
 
#Give it a nice Style so it stands out 
$range.Style = 'Title' 

#Increment row for next set of data 
$row++;$row++ 
 
#Save the initial row so it can be used later to create a border 
#Counter variable for rows 
$intRow = $row 
$xlOpenXMLWorkbook=[int]51


#Create Headers for Data
$MainSheet.Cells.Item($intRow,1)  ="Distributed Application Name" 
$MainSheet.Cells.Item($intRow,2)  ="ManagementPack" 
$MainSheet.Cells.Item($intRow,3)  ="SCOM Object ID" 

for ($col = 1; $col –le 3; $col++) 
     { 
          $MainSheet.Cells.Item($intRow,$col).Font.Bold = $True 
          $MainSheet.Cells.Item($intRow,$col).Font.ColorIndex = 41 
     } 
 
$intRow++ 

#get Distributed Applications
$das = Get-SCOMClass -ComputerName $scom -Credential $credential -displayname 'user created distributed application' | get-scomclassinstance -ComputerName $scom -Credential $credential



#Export DAs with Management Pack, Displayname and SCOM ObjectID (SCOM Object ID can be used in Dashboards)
$Daobjects = get-scomclass -ComputerName $scom -Credential $credential -displayname $das.DisplayName | Select-Object displayname, ManagementPackName,ID

#create new worksheet for each discovered Distributed Application
#add each distributed application to the main sheet with its Display name, ManagementPack and SCOM Object ID

foreach ($object in $daobjects){ #Begin work on main sheet
        $name = $object.displayname
        $mp = $object.managementpackname
        $id = $object.id
        $guid = $id.guid

        #Work on the main sheet that contains all distributed applications
        #adds the Distributed Application Name, Management Pack and Object ID to the Main Sheet
        $Mainsheet.Cells.Item($intRow, 1) = $name 
        $Mainsheet.Cells.Item($intRow, 2) = $mp
        $Mainsheet.Cells.Item($intRow, 3) = $guid

        $intRow = $intRow + 1 
        $MainSheet.UsedRange.EntireColumn.AutoFit()
        
        }#end work on main sheet

foreach ($daobject in $daobjects){  #Begin each individual worksheet.
        $daobjectname = $daobject.displayname
        $daobjectmp = $daobject.managementpackname
        $daobjectid = $daobject.id
        $daobjectguid = $id.guid

        #Checks the Length of the DA Name, truncates the Distributed Application Name to 31 characters if it is longer. Excel Sheet names are limited to 31 characters
        if($daobjectname.length -gt 31){$daobjectname = $daobjectname.remove(31)}

        #Checks for forward and backslashes, replace Slashes with Hyphen as Excel will not accept slashes in Sheet Names
        if($daobjectname -match '[\\/]'){
        $daobjectname = $daobjectname -replace $pattern, '-'}


        #work on each individual sheet for each Distributed Application
        #create worksheets for each Distributed Application
        $worksheet = $Excel.Worksheets.add()
        $worksheet.Name = "$daobjectname"

        #Create a Title for the Distributed Applicatoin
        $row = 1 
        $Column = 1 
        $worksheet.Cells.Item($row,$column)= "$daobjectName"

        $range = $worksheet.Range("a1","s2") 
        $range.Merge() | Out-Null 
        $range.VerticalAlignment = -4160 
 
        #Give it a nice Style so it stands out 
        $range.Style = 'Title' 

        #Increment row for next set of data 
        $row++;$row++ 
 
        #Counter variable for rows 
        $intRow = $row 

        #Create Headers for Data
        $workSheet.Cells.Item($intRow,1)  ="Distributed Application Component" 
        $workSheet.Cells.Item($intRow,2)  ="Component Object ID" 
        $workSheet.Cells.Item($intRow,3)  ="SCOM Group" 
        $workSheet.Cells.Item($intRow,4)  ="SCOM Group ObjectID" 
        $workSheet.Cells.Item($intRow,5)  ="Component Members" 
        $workSheet.Cells.Item($intRow,6)  ="SCOM Object ID" 

        for ($col = 1; $col –le 6; $col++) 
             { 
                  $WorkSheet.Cells.Item($intRow,$col).Font.Bold = $True 
                  $WorkSheet.Cells.Item($intRow,$col).Font.ColorIndex = 41 
             } 
        #increment row for the next data
        $intRow++ 

        #get Distributed Application Components Using GetRelatedMonitoringObjects()
        $Dacomps = get-scomclass -ComputerName $scom -Credential $credential -DisplayName $daobjectname | Get-SCOMClassInstance -ComputerName $scom -Credential $credential

        $components = $dacomps.GetRelatedMonitoringObjects()
            #Begin Foreach Distributed Application Component Loop
            foreach($component in $components)
            {
                  $name = $component.displayname
                  $id = $component.id
                  $guid = $id.guid

                  #add Distributed Application Component Name and SCOM Object ID to the worksheet
                  $worksheet.Cells.Item($intRow, 1) = $name 
                  $worksheet.Cells.Item($intRow, 2) = $guid
              


                  $objects = $component.GetRelatedMonitoringObjects() | select-object displayname, id

                  #Begin foreach distributed application component object loop
                  #loop through each object contained in Distributed Application Componet, if group get group members
                  foreach($object in $objects){

                        $objectname = $object.DisplayName
                        $objectid = $object.id
                        $objectguid = $objectid.guid
                   
                        #check if Object is a SCOM Group, if SCOM group get its group members, otherwise add component objects to spreadsheet
                        $verify = get-scomgroup -ComputerName $scom -Credential $credential -DisplayName $objectname
                  
                        If($verify -ne $null){
                        $groupdisplayname = $verify.displayname
                        $grpid = $verify.Id
                        $grpguid = $grpid.guid

                        #get SCOM Group Members
                        $objectgrps = $verify.GetRelatedMonitoringObjects()
                            #get scom group members and add to spreadsheet
                            foreach($objectgrp in $objectgrps){  #begin for each group
                                    $objectgrpname = $objectgrp.displayname
                                    $objectgrpID = $objectgrp.id
                                    $objectgrpguid = $objectgrpID.guid
                                         
                                    #add SCOM Group Name and ObjectID to worksheet
                                    #add SCOM Object Name and ID to worksheet
                                    $worksheet.Cells.Item($intRow, 3) = $groupdisplayname
                                    $worksheet.Cells.Item($intRow, 4) = $grpguid
                                    $worksheet.Cells.Item($intRow, 5) = $objectgrpname
                                    $worksheet.Cells.Item($intRow, 6) = $objectgrpguid
                                    $intRow++ #add new row
                                    } #End for each group

                        } Else{
                            #add SCOM Group Name and ObjectID to worksheet as N/A
                            #add SCOM Object Name and ID to worksheet
                            $worksheet.Cells.Item($intRow, 3) = "N/A" 
                            $worksheet.Cells.Item($intRow, 4) = "N/A"
                            $worksheet.Cells.Item($intRow, 5) = $objectname 
                            $worksheet.Cells.Item($intRow, 6) = $objectguid
                            $intRow++
                            } #End Else
                             
                        $workSheet.UsedRange.EntireColumn.AutoFit()


                        }#end Foreach Distributed Application Component object Loop



                }#end foreach Distributed Application Component Loop

       
       
       
} #end foreach loop of Distributed Applications



#Save file to specified location if SaveandClose is True
if($saveandclose -eq $true)
{

if (!(Test-Path -path "$path")) #create it if not existing 
  { 
  New-Item "$path" -type directory | out-null }


$file = "$path$filename.xlsx" 
if (test-path $file ) { remove-item $file } #delete the file if it already exists 

$Excel.SaveAs($file, $xlOpenXMLWorkbook) #save as an XML Workbook (xslx) 
$Excel.Saved = $True
$Excel.Close() 
$doc.quit()

}

#cleanup COM Objects
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($mainsheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc)

Remove-Variable -Name mainsheet
Remove-Variable -Name worksheet
Remove-Variable -Name excel
Remove-Variable -Name doc