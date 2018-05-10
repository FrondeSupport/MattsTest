Param(
 [Parameter(Mandatory=$true)]
 [string] $server,
 [Parameter(Mandatory=$true)]
 [int] $ErrorLimit,
  [Parameter()]
 [string] $WindowsS3BucketCSV,
  [Parameter()]
 [string] $WindowsS3Bucket,
   [Parameter()]
 [string] $WindowsS3BucketRegion,
 [switch] $Download
)


#The below parameters are all passed into the script when called. no need to set them here, but here are the uses:
##Set the ErrorLimit to a value if you want to stop the updates for a server on x number exceptions
#$ErrorLimit = 9999
##Set Download to -Download to enable automatic downloading of undownlaoded updates
#$Download=''
##$Download='-Download'
##The below controls where the logs, csv reports are contained
#$WindowsS3Bucket="windows-update-logging"
#$WindowsS3BucketCSV="windows-update-reports"
#$WindowsS3BucketRegion="ap-northeast-1"


$DATE=Get-Date -Format g
$DATELOG=Get-Date -Format d
$DATELOG=$DATELOG -replace "/",""

#Temporary locations for the log files before they go to S3.
$ReportCSV="C:\WindowsUpdateReport.$DATELOG.csv"
$LogFile="C:\WindowsUpdate.$DATELOG.log"



Function FrondeUpdate ($server, $Download, $ErrorLimit) {
    $LogString=''
    $intcount = 0
	FormatOutput "Starting script on server..."
	$Criteria = "IsInstalled=0 and Type='Software'"
	$Searcher = New-Object -ComObject Microsoft.Update.Searcher
	FormatOutput "About to search for Software updates that aren't installed.."
	try {
		$SearchResult = $Searcher.Search($Criteria).Updates
		FormatOutput "Searching for Software updates that aren't installed completed.."
		if ($SearchResult.Count -eq 0) {
			FormatOutput "There are no applicable updates."
		}
		else {
			$Session = New-Object -ComObject Microsoft.Update.Session
            if ($Download){
                FormatOutput "Downloading enabled. Starting download process."
			    $Downloader = $Session.CreateUpdateDownloader()
			    $Downloader.Updates = $SearchResult
			    $Downloader.Download() | out-null
                FormatOutput "Downloading finished. Moving onto installation."
            }
            else
            {
                FormatOutput "Downloading not enabled. Moving onto installation."
            }
			$Installer = New-Object -ComObject Microsoft.Update.Installer
            $UpdatesCollection = New-Object -ComObject Microsoft.Update.UpdateColl
            $UpdateCounter=0
            Foreach ($Update in $SearchResult)
            {
            try {
                $UpdateDownloaded = $Update.IsDownloaded
                $UpdateTitle = $Update.Title
                if ($UpdateDownloaded -eq "True" -and $UpdateTitle -ne ""){
                    $UpdateCounter++
                    $UpdatesCollection.Add($Update) | out-null
			        $Installer.Updates = $UpdatesCollection
			        $Result = $Installer.Install()
				    $HexCode = '{0:x4}' -f $Result.HResult
                    if ($Result.HResult -eq "0000"){
                        $LogString = $LogString + "SUCCESS: Update '$UpdateTitle' was installed on server $server. HResult code was $HexCode.`r`n"
                    }
                    else
                    {
                        $LogStringError = $LogStringError + "ERROR: There were issues with the install of update '$UpdateTitle' on server $server. HResult code was $HexCode.`r`n"
                        $intcount++
                    }
                    $UpdatesCollection.Clear()
                    if ( $intcount -ge $ErrorLimit ) { FormatOutput "ERROR - Error Limit breached. Exiting any further updates"; exit;}
                    }
                else{
                    FormatOutput "WARNING: Update '$UpdateTitle' was found but it's not downloaded. It will be skipped."
                    $LogString = $LogString + "WARNING: Update '$UpdateTitle' was found but it's not downloaded. It will be skipped..`r`n"

                }
                }
           catch {
                 $intcount++
                 $ErrorMessage = $_.Exception.Message
                 $FailedItem = $_.Exception.ItemName
		         FormatOutput "There was an exception - $FailedItem, $ErrorMessage"
        }

            }
        FormatOutput "There were $intcount errors installing on this server. There were $UpdateCounter updates installed"
        If ($Result.rebootRequired) {
            FormatOutput "RebootRequired"
            }
        else
            {
            FormatOutput "RebootNotRequired"
            }
        }
	}
	catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
		FormatOutput "There was an exception - $FailedItem, $ErrorMessage"

	}
    Add-Content $LogFile "$DATE - $server - Windows updates installed"
    Add-Content $LogFile "$LogString"
    Add-Content $LogFile $LogStringError
    if($S3Available){
    FormatOutput "The windows update log has been copied to a timestamped file in S3 bucket $S3. The logfile is called WindowsUpdate_$server.$DATELOG.log."
    Write-S3Object -BucketName $WindowsS3Bucket -File $LogFile -Key WindowsUpdate_$server.$DATELOG.log -CannedACLName bucket-owner-read -Region $WindowsS3BucketRegion
    }
    if ( $intcount -ne 0 ){
        FormatOutput "The errors are included below:`r`n$LogStringError"
		}
        FormatOutput "The full installation logs are included below:`r`n$LogString"
    return
}

function FormatOutput ($message){
    $DATE=Get-Date -Format g
    Write-Output "[$DATE] $server - $message"

    return
}

Function Get-MSHotfix
{
    $outputs = Invoke-Expression "wmic qfe list"
    $outputs = $outputs[1..($outputs.length)]


    foreach ($output in $Outputs) {
        if ($output) {
            $output = $output -replace 'y U','y-U'
            $output = $output -replace 'NT A','NT-A'
            $output = $output -replace '\s+',' '
            $parts = $output -split ' '
            if ($parts[5] -like "*/*/*") {
                $Dateis = [datetime]::ParseExact($parts[5], '%M/%d/yyyy',[Globalization.cultureinfo]::GetCultureInfo("en-US").DateTimeFormat)
            } elseif (($parts[5] -eq $null) -or ($parts[5] -eq ''))
            {
                $Dateis = [datetime]1700
            }

            else {
                $Dateis = get-date([DateTime][Convert]::ToInt64("$parts[5]", 16))-Format '%M/%d/yyyy'
            }
            New-Object -Type PSObject -Property @{
                KBArticle = [string]$parts[0]
                Computername = [string]$parts[1]
                Description = [string]$parts[2]
                FixComments = [string]$parts[6]
                HotFixID = [string]$parts[3]
                InstalledOn = Get-Date($Dateis)-format "dddd d MMMM yyyy"
                InstalledBy = [string]$parts[4]
                InstallDate = [string]$parts[7]
                Name = [string]$parts[8]
                ServicePackInEffect = [string]$parts[9]
                Status = [string]$parts[10]
            }
        }
    }
}


if ((Test-Path variable:global:WindowsS3BucketCSV) -And (Test-Path variable:global:WindowsS3Bucket) -And (Test-Path variable:global:WindowsS3BucketRegion)){
    $S3Available = $false
    FormatOutput "Disabling push to S3 bucket as one of the required parameters does not exist"
}

if (Test-Path variable:global:WindowsS3BucketCSV){
    $S3Available = $false
    FormatOutput "1"
}
if (Test-Path variable:global:WindowsS3Bucket){
    $S3Available = $false
    FormatOutput "2"
}
if (Test-Path variable:global:WindowsS3BucketRegion){
    $S3Available = $false
    FormatOutput "3"
}






if ($S3Available){
    if (Get-Module -ListAvailable -Name AWSPowerShell) {
        if (!(Get-Module AWSPowerShell)){
        Import-Module AWSPowerShell
        $S3Available = $true
        }
    } else {
            FormatOutput "AWS PowerShell is not available. This script will not copy files to S3."
            $S3Available = $false
    }
}


FrondeUpdate $server $Download $ErrorLimit


FormatOutput "Patching server completed. Now proceeding to run Windows update report"


Get-MSHotfix|Where-Object {$_.Installedon -gt ((Get-Date).Adddays(-90))}|Select-Object -Property Computername, KBArticle,InstalledOn, HotFixID, Description, InstalledBy| Export-CSV  $ReportCSV
if($S3Available){
    FormatOutput "Uploading CSV report to S3 Bucket - $WindowsS3BucketCSV"
    Write-S3Object -BucketName $WindowsS3BucketCSV -File $ReportCSV -Key "$DATELOG\WindowsUpdateReport_$server.$DATELOG.csv" -CannedACLName bucket-owner-read -Region $WindowsS3BucketRegion
}