<#
	.SYNOPSIS
		Generate a CMPG report for a NetApp SYSTEM_ID based on online autosupports.
	.DESCRIPTION
		This script gathers all the STATS files from the PERFORMANCE autosupports that have been send to NetApp for a specific systemid & time-range.
		Then it will run the offline CMPG tool to generate a report
	.PARAMETER systemid
		Please specify the SystemID for the NetApp you would like to generate the Report for.
	.PARAMETER outpath
		Please specify the path where you would like to save the STATS files 
	.PARAMETER startdate
		Please specify the StartDate in your own format (either DD/MM/YYYY or MM/DD/YYYY). If you don't specify anything it will take the current date -3 months
	.PARAMETER enddate
		Please specify the StartDate in your own format (either DD/MM/YYYY or MM/DD/YYYY). If you don't specify anything it will take the current date
	.PARAMETER username
		Please specify the username, if you won't specify it as a paramater you will be prompted for it
	.PARAMETER password
		Please specify the password, if you won't specify it as a paramater you will be prompted for it
    .PARAMETER nodownload
		Please specify if you want to download all the stats files, if you won't specify it as a paramater we will download everything.
	.PARAMETER reuse_asupfile
		Please specify if you would like to reuse an existing asup datatable, if you won't specify it as a paramater we will generate a new one.
	.EXAMPLE
		.\generate_cmpg.ps1 -systemid <id> -outpath <path>
	.Notes
		.Author 
			Boris Aelen
#>
 
 [cmdletbinding()]
param (
	[Parameter(Mandatory=$true)]	[string]$systemid,
	[Parameter(Mandatory=$true)]	[string]$outpath,
	[string]$password,
	[Datetime]$EndDate=(Get-Date),	
    [Datetime]$StartDate=$EndDate.AddMonths(-3),
	[string]$username,
    [boolean]$nodownload=0,
    [boolean]$reuse_asupfile=0
)
$ErrorActionPreference= 'Stop'


#Configure Variables
$login_url="https://signin.netapp.com/obrareq.cgi?wh%3Dutp.corp.netapp.com%20wu%3D%2FHome%20wo%3D1%20rh%3Dhttp%3A%2F%2Fsmartsolve.netapp.com%20ru%3D%252FHome"
$downloadurl = "https://smartsolve.netapp.com/search/asup?exportAll=1&loadAll=1&sysid="+$systemid
$outpath2=$outpath+"\"+$systemid
$excelfile = $outpath2 + "\smartsolve_output_"+$systemid+".xls"
$asupsfile = $outpath2 + "\asups_file.xls"
$7mode = 0

cls
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""


#Validating & Creating the outpath2
if (-not (Test-Path $outpath)) {Write-host "Path '$outpath' does not exist."; exit}
if (!(Test-Path $outpath2)) { $r = mkdir $outpath2 }

if (!$nodownload){
    #Validating & requesting the credentials
    if ($username -eq "") { $username = Read-Host 'username' } 
    if ($password -eq "") { $securepassword = Read-Host 'password' -AsSecureString } else { $securepassword = $password | ConvertTo-SecureString -AsPlainText -Force }
    $Credentials= New-Object -typename System.Management.Automation.PSCredential -argumentlist  $username, $securepassword
    $PlainPassword = $Credentials.GetNetworkCredential().Password

    #Login to NetApp
    Write-host -NoNewline  "Logging in to the NetApp Smartsolve site... " 
    $r=Invoke-WebRequest $login_url -SessionVariable session 
    $form = $r.Forms[0]
    $form.Fields["user"] = $username
    $form.Fields["password"] =  $PlainPassword
    $ErrorActionPreference= 'silentlycontinue'
    $res=Invoke-WebRequest -Uri ($login_url + $form.Action) -WebSession $session -Method POST -Body $form.Fields 
    if ($res -eq $null) { Write-host  "Success" } else { Write-host "FAILED"; exit }
    $ErrorActionPreference= 'Stop'
}

if (!$reuse_asupfile){
  
    #Get the Excel File 
    Write-Host -NoNewline  "Downloading the list of all the ASUPS for $systemid... "
    write-host $downloadurl 
    $r=Invoke-WebRequest -Uri $downloadurl -WebSession $session -Outfile $excelfile 
    Write-Host "Success"

    #Load the Excel File

    if ($excelfile -eq "") { Write-host "Please provide path to the Excel file"; Exit}
    if (-not (Test-Path $excelfile)) {Write-host "Path '$excelfile' does not exist."; exit}
    
    $excel = New-Object -com "Excel.Application"
    $excel.Visible = $false

    $newci = [System.Globalization.CultureInfo]"en-US"
    [system.threading.Thread]::CurrentThread.CurrentCulture = $newci
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $newci

    
    $workbook = $excel.workbooks.open($excelfile)
    $sheet = $workbook.ActiveSheet
    if (-not $sheet) { Write-host "Unable to open worksheet $WorksheetName";  exit }
    $columns = $sheet.UsedRange.Columns.Count
    $lines = $sheet.UsedRange.Rows.Count
    $sheetName = $sheet.Name

    #Create a DataTable for all the relevant Information  	
    $tabName = "All ASUPS"
    $asups = New-Object system.Data.DataTable "$tabName"
    #Define Columns
    $Date = New-Object system.Data.DataColumn Date,([datetime])
    $URL = New-Object system.Data.DataColumn URL,([string])
    $Title = New-Object system.Data.DataColumn Title,([string])
    $Release = New-Object system.Data.DataColumn Release,([string])
    #Add the Columns
    $asups.columns.add($Date)
    $asups.columns.add($URL)
    $asups.columns.add($Title)
    $asups.columns.add($Release)

    Write-Host -NoNewline  "Importing the EXCEL file... " 
    #Fill the DataTable
    for ($line = 2; $line -le $lines; $line ++) {
	    $row = $asups.NewRow()
	    $row.Date = $sheet.Cells.Item.Invoke($line, 1).Value2.substring(0,21)
	    $row.Title = $sheet.Cells.Item.Invoke($line, 2).Value2 
	    $row.URL = $sheet.Cells.Item.Invoke($line, 1).Hyperlinks | select -expand Name		
        $row.Release = $sheet.Cells.Item.Invoke($line, 8).Value2 
        if ($row.Date -ge $StartDate -and $row.Date -le $EndDate -and $row.Title -like "*PERFORMANCE*") {	
	        $asups.Rows.Add($row)
        }
	
	    $percents = [math]::round((($line/$lines) * 100), 0)
        Write-Progress -Activity:"Importing from Excel file $excelfile" -Status:"Imported $line of total $lines lines ($percents%)" -PercentComplete:$percents
    }
    #Close the Excel File  
    $workbook.Close()
    $excel.Quit()  
    Write-Host "SUCCESS" 
    $asups | Export-Clixml $asupsfile

    $newci = [System.Globalization.CultureInfo]"nl-NL"
    [system.threading.Thread]::CurrentThread.CurrentCulture = $newci
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $newci
}

if (!$nodownload){
    if ($reuse_asupfile){
        #Create a DataTable for all the relevant Information  	
        #$tabName = "All ASUPS"
        #$asups = New-Object system.Data.DataTable "$tabName"
        $asups = Import-Clixml $asupsfile
    }
    Write-Host -NoNewline "Downloading the STATS file from SmartSolve... " 
    $asups | foreach {
    	    $asupid = $_.URL.Split("=") | select -skip 1
        if ($_.Release.EndsWith("7-Mode")) {
            $7mode = 1
            $stats_data_url = "https://smartsolve.netapp.com/cdv/rawsectiondata?section=CM-HOURLY-STATS.GZ&asupids="+$asupid+"&tzoffset=120&timezone=WEDT"
		    $stats_data_file = $outpath2 + "\" + ("{0:yyyyMMdd}" -f [datetime]$_.Date) + "_hourly_stats.gz"        
	        Write-Progress -Activity:"Downloading STATS files from SmartSolve" -Status:"Downloading $("{0:dd-MM-yyyy}" -f [datetime]$_.Date)..." 
            if (!(Test-Path $stats_data_file)){ 
                write-host "download " + $stats_data_url
    		    $r=Invoke-WebRequest -Uri $stats_data_url -WebSession $session -Outfile $stats_data_file 
            } else { Write-host "file exists "+$stats_data_file }
        } else { 
            $7mode = 0
            $stats_data_url = "https://smartsolve.netapp.com/cdv/downloadcdvsection?section=CM-STATS-HOURLY-DATA.XML&asupids="+$asupid+"&formatType=plain"
		    $stats_info_url = "https://smartsolve.netapp.com/cdv/downloadcdvsection?section=CM-STATS-HOURLY-INFO.XML&asupids="+$asupid+"&formatType=plain"
		    $stats_data_file = $outpath2 + "\" + ("{0:yyyyMMdd}" -f [datetime]$_.Date) + "_hourly_data.xml"  
		    $stats_info_file = $outpath2 + "\" + ("{0:yyyyMMdd}" -f [datetime]$_.Date) + "_hourly_info.xml" 
	        Write-Progress -Activity:"Downloading STATS files from SmartSolve" -Status:"Downloading $("{0:dd-MM-yyyy}" -f [datetime]$_.Date)..." 
            if (!(Test-Path $stats_data_file)){ 
    		    $r=Invoke-WebRequest -Uri $stats_data_url -WebSession $session -Outfile $stats_data_file 
            } else { Write-host "file exists "+$stats_data_file }
            if (!(Test-Path $stats_info_file)){ 
	    	    $r=Invoke-WebRequest -Uri $stats_info_url -WebSession $session -Outfile $stats_info_file
            } else { Write-host "file exists "+$stats_info_file }
        }
    }
	Write-Host "SUCCESS" 
    Write-Host "Downloaded all the STATS file for $systemid into $outpath2. Opening directory now."
    ii $outpath2
}
	

#Run CPMG
if ($7mode){
    $cmpgcmd = "c:\cmpg\cmpg\cmpg.exe -f "+ $outpath2 + "\*.gz -details"
    #write-host $cmpgcmd
    Invoke-Expression $cmpgcmd
    $mvcmd = "mv "+$outpath2+"\*-perf.xlsx " + $outpath+"\"+$systemid+"-perf.xlsx"
    Invoke-Expression $mvcmd

    $cmpgcmd = "c:\cmpg\cmpg\cmpg.exe -f "+$outpath2+"\*.gz  -xml C:\cmpg\cmpg\trend.xml"
    #write-host $cmpgcmd
    Invoke-Expression $cmpgcmd
    $mvcmd = "mv "+$outpath2+"\*-trending.xlsx "+$outpath+"\"+$systemid+"-trending.xlsx"
    Invoke-Expression $mvcmd

} else {
    $cmpgcmd = "c:\cmpg\cmpg\cmpg.exe -f "+ $outpath2 + "\*.xml -details"
    #write-host $cmpgcmd
    Invoke-Expression $cmpgcmd
    $mvcmd = "mv "+$outpath2+"\*-perf.xlsx " + $outpath+"\"+$systemid+"-perf.xlsx"
    Invoke-Expression $mvcmd

    $cmpgcmd = "c:\cmpg\cmpg\cmpg.exe -f "+$outpath2+"\*.xml  -xml C:\cmpg\cmpg\trend.xml"
    #write-host $cmpgcmd
    Invoke-Expression $cmpgcmd
    $mvcmd = "mv "+$outpath2+"\*-trending.xlsx "+$outpath+"\"+$systemid+"-trending.xlsx"
    Invoke-Expression $mvcmd
}
