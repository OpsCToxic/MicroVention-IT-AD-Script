<#
Getting Data
- Connect to Dataverse Asset Table
- Get Assets
    - Filters
        Asset Status = Assigned
        Model Vendor = Dell
Foreach Service Tag
    - Get Computer Object details from AD within MicroVention OU
        - Operating System Version
    - SWITCH If(OS Version eq ...){
        Select appropriate OS version from Lookup table in Operating System Dataverse table
        Write OS Version to Asset record on Dataverse
    }
#>


#Import-Module ActiveDirectory
#Install-Module -Name Microsoft.Xrm.Tooling.CrmConnector.PowerShell
#Import-Module -Name Microsoft.Xrm.Data.PowerShell


# ============ Constants to change ============
$fetchxml = "C:\Users\mv.lansweeper\Documents\PowerShell Scripts\FetchXMLs\Fetch_ServiceTagsOS.xml"
# =============================================


$date = get-date -format "MM-dd-yyyy.HH:mm:ss" | ForEach-Object { $_ -replace ":", "." }
$filename = 'C:\Users\mv.lansweeper\Documents\PowerShell Scripts\Dataverse Update DV with OS Logs\' + $date + '.txt'


# Create file for logging content as it is queried from Lansweeper
Start-Transcript -Path $filename


# ============ Login to MS CRM ============
#Get credentials for connection
#$credential = Get-Credential
#Connect-ADAccount -Credential $credential


#$Cred = Get-Credential
#$CRMConn = Get-CrmConnection -Credential $Cred -DeploymentRegion NorthAmerica -OnlineType Office365 -OrganizationName org20cecdc4
$Cred = Get-StoredCredential -Target "Dataverse"
$CRMConn = Get-CrmConnection -Credential $Cred -DeploymentRegion NorthAmerica -OnlineType Office365 -OrganizationName org20cecdc4


# ============ Fetch data ============
[string]$fetchXmlStr = Get-Content -Path $fetchxml


$list = New-Object System.Collections.ArrayList


# Be careful, NOT zero!
$pageNumber = 1
$pageCookie = ''
$nextPage = $true


$StartDate1=Get-Date


# This will query the content from Lansweeper by storing all computer object records from Lansweeperin the $list ArrayList collection
while($nextPage)
{    
    if ($pageNumber -eq 1) {
        $result = Get-CrmRecordsByFetch -conn $CRMConn -Fetch $fetchXmlStr
    }
    else {
        $result = Get-CrmRecordsByFetch -conn $CRMConn -Fetch $fetchXmlStr -PageNumber $pageNumber -PageCookie $pageCookie
    }
   
    $EndDate1=Get-Date
    $ts1 = New-TimeSpan –Start $StartDate1 –End $EndDate1
 
    $list.AddRange($result.CrmRecords)
 
    Write-Host "Fetched $($list.Count) records in $($ts1.TotalSeconds) sec"
     
    $pageNumber = $pageNumber + 1
    $pageCookie = $result.PagingCookie
    $nextPage = $result.NextPage
}


# $list is a collection of pages of queried content of device information from Lansweeper
# $tag is the current serial number being evaluated in the loop from $list collection
foreach ($tag in $list) {


    # Store the serial number property from each object in $list and assign it to $serivceTag.
    $serviceTag = $tag.mvi_serialnumber
    # Display serial number to console
    Write-Host "Service Tag:" $serviceTag
    # Store Asset ID of each object and store in $assetID
    $assetID = $tag.mvi_assetid
    # Store Asset ID of each object and store in $assetID
    $OSLookupID = $tag.mvi_operatingsysteminstalled
    # Display OS lookup ID to console
    Write-Host "OS Lookup Field:" $OSLookupID


    #Build object for OS Lookup Field
    $lookupOSObject = New-Object -TypeName Microsoft.Xrm.Sdk.EntityReference;
    $lookupOSObject.LogicalName = "mvi_operatingsystem";


    # Retrieves the Operating System field from Active Directory by filtering the table of computers by $serviceTag and expanding the OperatingSystemVersion property of the computer object
    $operatingSystemBuild = Get-ADComputer -Filter {name -eq $serviceTag} -SearchBase "OU=Microvention,DC=us,DC=terumo,DC=com" -Properties * | Select-Object -ExpandProperty OperatingSystemVersion
    Write-Host "AD Build Number:" $operatingSystemBuild


    # Compares the build of the retrieved operating system and assigns it with a unique ID that correlates to the Operating System lookup table. This will allow us to reference the Operating System
    # Table from the Asset table by assigning a new column in the Asset table with these unique IDs.
    switch ($operatingSystemBuild) {
      # Windows 7 Enterprise NT 6.1
        "6.1 (7601)" {
            $lookupOSObject.Id = "f13444c8-ec0a-ee11-8f6e-000d3a370dc4"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 1703
        "10.0 (15063)" {
            $lookupOSObject.Id = "097539c8-ed0a-ee11-8f6e-000d3a370dc4"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 1709
        "10.0 (16299)" {
            $lookupOSObject.Id = "4ff4b684-ed0a-ee11-8f6e-000d3a370dc4"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 1809
        "10.0 (17763)" {
            $lookupOSObject.Id = "daf33644-2a77-ed11-81ab-000d3a370dc4"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 1903
        "10.0 (18362)" {
            $lookupOSObject.Id = "e283e151-4d06-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 1909
        "10.0 (18363)" {
            $lookupOSObject.Id = "2552ff7a-4d06-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 2004
        "10.0 (19041)" {
            $lookupOSObject.Id = "726407c7-4d06-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 20H2
        "10.0 (19042)" {
            $lookupOSObject.Id = "976406db-4d06-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 21H1
        "10.0 (19043)" {
            $lookupOSObject.Id = "0774dd02-4e06-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 21H2
        "10.0 (19044)" {
            $lookupOSObject.Id = "70791e49-4e06-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 10 Enterprise 22H2
        "10.0 (19045)" {
            $lookupOSObject.Id = "3bf04e2e-4e06-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 11 Enterprise 21H2
        "10.0 (22000)" {
            $lookupOSObject.Id = "c9b95584-ec0a-ee11-8f6e-000d3a370dc4"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
      # Windows 11 Enterprise 22H2
        "10.0 (22621)" {
            $lookupOSObject.Id = "3b948ae0-cd0a-ee11-8f6e-000d3a3706fd"
            Set-CrmRecord -conn $CRMConn -EntityLogicalName mvi_asset -Id $assetID -Fields @{"mvi_operatingsysteminstalled"=[Microsoft.Xrm.Sdk.EntityReference] $lookupOSObject}
          }
        Default {
            Write-Host "Operating System Build Number Not Found"
        }


    }




   
   
}


Stop-Transcript

