<#
## Host Audit Function - Visual Output ########################
written by Chris Renshaw
Created 02/02/2016

Note - before you can run this script you will need to update 
certain settings in it to meet your infrastructure - see the 
variables section

Dependency Needed - create C:\PowerShell and place this script and 
Dependency-JoinObject.ps1 into this folder. JoinObject script was 
developed by Microsoft and included here for convenience. I made no
modifications to that script at all.

## Outline of script ##########################################

* Scans specified PCs or OU ($Hosts) for ping connectivity - 
  email notification for any not answering
* Scans $Hosts for utilization of CPU, RAM, and HDD space 
  against pre-defined limits - email notification for any 
  crossing the limit
* Logs CSV files for both disk and specs of machine

## Considerations (To Do) ######################################

* Modify output variable post-function to easily give results 
  upon inquiry ($MasterVariable[0].RAMFree (or similar) results in amount of 
  RAM free for the first listed host)

#> #############################################################
#TAGS Query,Computers,OU,Host,Audit,Hardware,Comp,Spec,Specs
#UNIVERSAL

function Host-Audit {

    function HostSpecs {
        clear

        # Variables to modify ####################################################################################
        $to = "recipient@email.com"
        $from = "sender@email.com"
        $smtpserver = "mailserver.email.com"
        $WorkingDirectory = "C:\PowerShell"
        $ReportingPath = "$WorkingDirectory\Output\Reports"
        $TempPath = "$ReportingPath\Temp"
        $CPULimit = "85" #greater than % usage
        $RAMLimit = "85" #greater than % usage
        $HDDLimit = "15" #less than % free

        #Choose one or the other option for the $Hosts Variable (add or remove the "#") 
            #$Hosts = Get-ADComputer -Filter {Name -like "*Host-or-other-search-term*"} -Searchbase "ou=Enter,ou=appropriate,ou=ServerOU,dc=Domain,dc=local"
            #$HostNameStamp = $Hosts.Name
        
            ## OR ##
        
            $Hosts = "HOST1","HOST2","HOST3"
            $HostNameStamp = $Hosts

        

        # Set Variables ###########################################################################################
        $HostOnline = @()
        $HostOffline = @()
        $HostReport = @()
        $PingHost = 'y'
        $DateStamp = Get-Date -UFormat "%m.%d.%Y"
        
        
        # Dependency Import #######################################################################################
        # Executes all files that begin with "Dependency" from the project's working directory
        cd $WorkingDirectory

        Foreach ($Dependency in (Get-ChildItem | Where {$_.Name -LIKE "Dependency*"})) {
            Invoke-Expression (". .\" + $Dependency.Name)
        }
        
        
        if (!(Test-Path $ReportingPath)) {
            New-Item $ReportingPath -type directory
        }

        cd $ReportingPath
        if (!(Test-Path $TempPath)) {
            New-Item $TempPath -type directory
        }

        # Introduction ############################################################################################
        Write-Host "~~~~~~~~~~~~~~~~~~" -ForegroundColor Magenta
        Write-Host "~~~ Host Audit ~~~" -ForegroundColor Magenta
        Write-Host "~~~~~~~~~~~~~~~~~~" -ForegroundColor Magenta
        Write-Host
        Write-Host "The following computers will be scanned:" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        $HostNameStamp
        Write-Host 
        Write-Host "Starting Connectivity Test..." -ForegroundColor Yellow

        # Ping Test ###############################################################################################
        foreach ($Computer in $HostNameStamp){
            If (Test-Connection $Computer -BufferSize 16 -Count 1 -Quiet)
                {Write-Host $Computer "is available over the network" -ForegroundColor Green}
            else
            {
                Write-Host $Computer "is NOT available on the network (not responding to PING)" -ForegroundColor Red
                $PingHost = 'n'
            }
            if ($PingHost -ne 'n'){$HostOnline += $Computer}
            if ($PingHost -eq 'n'){$HostOffline += $Computer}
        }
        if ($HostOffline.Count -gt 0){
            Write-Host
            Write-Host "Hosts that couldn't be contacted via PING have been removed from this list.  These include:" -ForegroundColor Yellow
            Write-Host 
            Write-Host $HostOffline -Foregroundcolor Red
            Write-Host
            Write-Host "An email alert has been sent to $to for troubleshooting" -ForegroundColor Yellow
            #Send an alert of failed ping
            Send-MailMessage -to $to -subject "PING Connectivity Error Report" -from $from -body "Warning, Host Connectivity is down. The following are not responding to PING. You may want to look into that. Please check on the following
            $HostOffline" -smtpserver $smtpserver
        }
        
        Write-Host
        Write-Host

        # Establish Session ######################################################################################
            if (!(Get-PSSession -ComputerName $HostOnline)) {
                Write-Host "No existing PSSession detected - establishing sessions" -ForegroundColor Yellow
                try {
                    $Sessions = New-PSSession -ComputerName $HostOnline
                }
                catch {
                    Write-Warning "Could not open WinRM on $HostOnline"
                }
            }

            else {
                Write-Host "Existing session detected.  No additional PSSession required." -ForegroundColor Green
                $Sessions = Get-PSSession
            }
            Write-Host
            Write-Host
            Write-Host "Beginning Hardware Scan. Please Wait..." -ForegroundColor Yellow


        # RAM Tests ###############################################################################################
        $RAM_Audit = @()

        $RAM_Audit  = Invoke-Command -Session $Sessions -ScriptBlock {
            param($RAMLimit)
            
            $RAM_Max = @()
            $RAM_Free = @()
            $RAM_Pct_Used = @()
            $Temp_RAM_Used = @()
                                                        
            $RAM_Max = [math]::round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
            $RAM_Free = [math]::round((Get-Counter '\Memory\Available MBytes').CounterSamples.CookedValue / 1024)
            $RAM_Utilized_Pct = [math]::round(($RAM_Free / $RAM_Max) * 100)

                if ($RAM_Utilized_Pct -ge $RAMLimit ){$RAM_AlertFlag = $TRUE} 
                else {$RAM_AlertFlag = $FALSE}

            $Temp_RAM_Used = New-Object –TypeName PSObject
            $Temp_RAM_Used | Add-Member –MemberType NoteProperty -Name RAM_Target –Value $env:COMPUTERNAME
            $Temp_RAM_Used | Add-Member –MemberType NoteProperty -Name RAM_Maximum –Value $RAM_Max
            $Temp_RAM_Used | Add-Member –MemberType NoteProperty -Name RAM_Available –Value $RAM_Free
            $Temp_RAM_Used | Add-Member -MemberType NoteProperty -Name RAM_Utilized_Pct -Value $RAM_Utilized_Pct 
            $Temp_RAM_Used | Add-Member –MemberType NoteProperty -Name RAM_Threshold –Value $RAMLimit
            $Temp_RAM_Used | Add-Member –MemberType NoteProperty -Name RAM_AlertFlag –Value $RAM_AlertFlag
            $Temp_RAM_Used

        } -Args $RAMLimit


        # CPU Tests ###############################################################################################
        
        Function Measure-CPU {
            $MeasurementCount = '4'    #Number of measurements to take within the measurement interval
            $MeasurementInterval = '4'   #Measurement interval (seconds)
            $Counters = ((Get-Counter '\processor(_total)\% processor time' -SampleInterval $MeasurementInterval -MaxSamples $MeasurementCount).CounterSamples.CookedValue)
            ForEach ($Counter in $Counters) {
                $Sum += $Counter -AS [decimal]
            }
            ([math]::Round(($Sum / $MeasurementCount),1)) # This is the return value

        }  # / Measure CPU

        $Output_CPU = Invoke-Command -Session $Sessions -ScriptBlock ${Function:Measure-CPU} 
                        
        $CPU_Audit = @()

        foreach ($CPU_Output in $Output_CPU) {
            $Temp_CPU_Used = New-Object –TypeName PSObject
            $Temp_CPU_Used | Add-Member –MemberType NoteProperty -Name CPU_Target –Value $CPU_Output.PSComputerName
            $Temp_CPU_Used | Add-Member –MemberType NoteProperty –Name CPU_Average –Value $CPU_Output
            $Temp_CPU_Used | Add-Member –MemberType NoteProperty –Name CPU_Threshold –Value $CPULimit
                           
            if ($CPU_Output -ge $CPULimit) {
                $Temp_CPU_Used | Add-Member –MemberType NoteProperty –Name CPU_AlertFlag -Value $TRUE
                $CPU_Audit += $Temp_CPU_Used 
                Remove-Variable Temp_CPU_Used
            }
                            
            else {
                $Temp_CPU_Used | Add-Member –MemberType NoteProperty –Name CPU_AlertFlag -Value $FALSE
                            
                $CPU_Audit += $Temp_CPU_Used 
                Remove-Variable Temp_CPU_Used
            }
        }


        # HDD Tests ###############################################################################################
        
        $HDD_Audit = Invoke-Command -Session $Sessions -ScriptBlock {
            param($HDDLimit)

            $DiskArray = (gwmi Win32_LogicalDisk -Filter "DriveType=3")
            $DiskList = @()
            $HDDArray = @()

            $HDDArray = New-Object -TypeName PSObject
            $HDDArray | Add-Member -MemberType NoteProperty -Name HDD_Target –Value $env:COMPUTERNAME
            $HDDArray | Add-Member –MemberType NoteProperty -Name HDD_DateStamp –Value $DateStamp 
            $HDDArray | Add-Member -MemberType NoteProperty -Name HDD_Disks –Value @()
            $HDDArray | Add-Member -MemberType NoteProperty -Name HDD_AlertFlag –Value $False

            ForEach ($Disk in $DiskArray) {
                $DiskName = $Disk.Name
                [int]$DiskFree = $Disk.FreeSpace / 1GB
                $DiskFree = "{0:N2}" -f $DiskFree
                [int]$DiskMax = $Disk.Size / 1GB
                $DiskMax = "{0:N2}" -f $DiskMax
                $DiskUsed = ($DiskMax - $DiskFree)
                [int]$DiskPercentFree = [math]::round(($DiskFree / $DiskMax) * 100)
                    if ($DiskPercentFree -ge $HDDLimit ) {$HDD_DiskAlertFlag = $False} 
                    else {
                    $HDD_DiskAlertFlag = $True
                    $HDDArray.HDD_AlertFlag = $True
                    }
                     
                $DiskObject = New-Object PSObject
                $DiskObject | Add-Member -MemberType NoteProperty -Name HDD_Target –Value $env:COMPUTERNAME
                $DiskObject | Add-Member -MemberType NoteProperty -name HDD_DiskName -Value $DiskName
                $DiskObject | Add-Member -MemberType NoteProperty -name HDD_DiskFree -Value $DiskFree
                $DiskObject | Add-Member –MemberType NoteProperty -Name HDD_DiskUsed –Value $DiskUsed
                $DiskObject | Add-Member –MemberType NoteProperty -Name HDD_DiskMax –Value $DiskMax
                $DiskObject | Add-Member –MemberType NoteProperty -Name HDD_DiskFreePct –Value $DiskPercentFree
                $DiskObject | Add-Member –MemberType NoteProperty -Name HDD_DiskThreshold –Value $HDDLimit
                $DiskObject | Add-Member –MemberType NoteProperty -Name HDD_DiskAlertFlag –Value $HDD_DiskAlertFlag
                $HDDArray.HDD_Disks += $DiskObject  

                
            } 
            $HDDArray
        }   -ArgumentList $HDDLimit



        # PREPARE $MASTER_OUTPUT object ###################################################   
        $Master_Output = @()       #Create empty array
        ForEach ($Target in $Hosts) { #Uses the $HOSTS input to create a custom object array called $MASTER_OUTPUT
            $Temp_Master_Output = New-Object –TypeName PSObject
            $Temp_Master_Output  | Add-Member –MemberType NoteProperty -Name Target –Value $Target
            $Master_Output += $Temp_Master_Output   
            Remove-Variable Temp_Master_Output
        }


        # JOIN LOGIC ######################################################################
        ### JOIN all results objecsts into $MASTER_OUTPUT object 
        # Utilizes JOIN-OBJECT.PS1 script developed by Microsoft
        $Collect_Spec_Output = Get-Variable | Where {$_.Name -LIKE "*_Audit"}   #Collects all results variables
        ForEach ($Output_Object in $Collect_Spec_Output) {
            $Custom_Properties = ($Output_Object.Value | gm -type NoteProperty).name   # Gathers the custom object's PROPERTY names
            $Join_Seam = $Custom_Properties | Where {$_ -LIKE "*Target*" }             #  Locates the property names $TESTNAME_Audit.TESTNAME_Target to use as a join seam 
            $Master_Output = Join-Object -Left $Master_Output -Right $Output_Object.Value -Where {$args[0].Target -eq $args[1].($Join_Seam)} -LeftProperties * -RightProperties $Custom_Properties
            # Runs an INNER JOIN on the TARGET column.  Combines all $RESULTS_TESTNAME variables into $MASTER_OUTPUT variable.
        } 
        

        # Convert results to readable output and error logging #############################
        $MasterObject = @()
        $Master_Output | ForEach-Object { 
            $MasterObject = New-Object PSObject
            $MasterObject | Add-Member –MemberType NoteProperty -Name Master_Target –Value $_.Target
            $MasterObject | Add-Member –MemberType NoteProperty -Name RAM_AlertFlag –Value $_.RAM_AlertFlag
            $MasterObject | Add-Member –MemberType NoteProperty –Name CPU_AlertFlag -Value $_.CPU_AlertFlag
            $MasterObject | Add-Member -MemberType NoteProperty -Name HDD_AlertFlag –Value $_.HDD_AlertFlag                           
            $MasterObject | Add-Member –MemberType NoteProperty -Name RAM_Maximum –Value $_.RAM_Maximum
            $MasterObject | Add-Member –MemberType NoteProperty -Name RAM_Available –Value $_.RAM_Available
            $MasterObject | Add-Member -MemberType NoteProperty -Name RAM_Utilized_Pct -Value $_.RAM_Utilized_Pct
            $MasterObject | Add-Member –MemberType NoteProperty -Name RAM_Threshold –Value $_.RAM_Threshold
            $MasterObject | Add-Member –MemberType NoteProperty –Name CPU_Average –Value $_.CPU_Average
            $MasterObject | Add-Member –MemberType NoteProperty –Name CPU_Threshold –Value $_.CPU_Threshold
            $MasterObject | Add-Member -MemberType NoteProperty -Name HDD_Disks –Value $_.HDD_Disks


            # Report Logged to CSV #########################################################
            $MasterTargetName = $MasterObject.Master_Target

            # Report for Specs (CPU and RAM) ###
            cd $ReportingPath

            $SpecFullPath = @()
            $SpecFullPath = "Host_Spec_Report_" + "$DateStamp" + ".CSV"
            $SpecConvert = ConvertTo-CSV -InputObject $MasterObject -NoTypeInformation
            Export-CSV "$SpecFullPath" -InputObject $MasterObject -Append -NoTypeInformation

            # Initial Report for HDD - Final report later in script ###
            cd $TempPath

            $HDD_DiskSpecs = $($MasterObject.HDD_Disks)
            $HDDFullPath = @()
            $HDDFullPath = "HDDSpecs_" + "$MasterTargetName" + "_" + "$DateStamp" + ".TXT"

            $HDD_DiskSpecs | Out-File $HDDFullPath
            

            # Error Reporting for Issues ###################################################
                cd $TempPath
                
                $RAM_Alerting = @()
                $CPU_Alerting = @()
                $HDD_Alerting = @()
                $RAM_Alerting = $($MasterObject.RAM_AlertFlag)
                $CPU_Alerting = $($MasterObject.CPU_AlertFlag)
                $HDD_Alerting = $($MasterObject.HDD_AlertFlag)

            # Error Reporting for RAM ######################################################
            if ($RAM_Alerting -eq $TRUE) {
                $RAM_Specs = New-Object PSObject
                $RAM_Specs | Add-Member –MemberType NoteProperty -Name RAM_Target –Value $MasterObject.Master_Target
                $RAM_Specs | Add-Member –MemberType NoteProperty -Name RAM_Maximum –Value $MasterObject.RAM_Maximum
                $RAM_Specs | Add-Member –MemberType NoteProperty -Name RAM_Available –Value $MasterObject.RAM_Available
                $RAM_Specs | Add-Member -MemberType NoteProperty -Name RAM_Utilized_Pct -Value $MasterObject.RAM_Utilized_Pct 
                $RAM_Specs | Add-Member –MemberType NoteProperty -Name RAM_Threshold –Value $MasterObject.RAM_Threshold
                $RAM_Specs | Add-Member –MemberType NoteProperty -Name RAM_AlertFlag –Value $MasterObject.RAM_AlertFlag
                
                $RAMAlertPath = @()
                $RAMAlertPath = "Alert_RAM_Report_" + "$DateStamp" + ".CSV"
                $RAMConvert = ConvertTo-CSV -InputObject $RAM_Specs -NoTypeInformation
                Export-CSV "$ReportingPath\$RAMAlertPath" -InputObject $RAM_Specs -Append -NoTypeInformation

            }
             

            # Error Reporting for CPU ######################################################
            if ($CPU_Alerting -eq $TRUE) {
                $CPU_Specs = New-Object PSObject
                $CPU_Specs | Add-Member –MemberType NoteProperty -Name CPU_Target –Value $MasterObject.Master_Target
                $CPU_Specs | Add-Member –MemberType NoteProperty –Name CPU_Average –Value $_.CPU_Average
                $CPU_Specs | Add-Member –MemberType NoteProperty –Name CPU_Threshold –Value $_.CPU_Threshold
                $CPU_Specs | Add-Member –MemberType NoteProperty –Name CPU_AlertFlag -Value $_.CPU_AlertFlag
                
                $CPUAlertPath = @()
                $CPUAlertPath = "Alert_CPU_Report_" + "$DateStamp" + ".CSV"
                $CPUConvert = ConvertTo-CSV -InputObject $CPU_Specs -NoTypeInformation
                Export-CSV "$ReportingPath\$CPUAlertPath" -InputObject $CPU_Specs -Append -NoTypeInformation

            }

            # Error Reporting for HDD ######################################################
            if ($HDD_Alerting -eq $TRUE) {
                $HDD_DiskSpecs = $($MasterObject.HDD_Disks)
                $HDDFullPath = @()
                $HDDFullPath = "Alert_HDD_" + "$MasterTargetName" + "_" + "$DateStamp" + ".TXT"
                $HDD_DiskSpecs | Out-File "$HDDFullPath"

            }   
            
        }

        
        # HDD Alert Consolidation and Report
        
        cd $TempPath
        $TempReportPull = @()

        # Remove blank lines
        $TempReportPull = Get-ChildItem | Where {$_.Name -LIKE "Alert_HDD*"}
        
        if ($TempReportPull -ne $NULL) {
            Get-Content $TempReportPull | where {$_.trim() -ne "" } > "HDD_Alert_Combined.txt"
            $HDD_Alert_Combined = "HDD_Alert_Combined.txt"
        
            Do { 
                # Trim away unneeded data from import
                Clear-Variable -Name ChunkTrimmed,ChunkTrimming,ChunkTrimming2,ChunkObject,GetNextChunk
                $ChunkTrimmed = @()
                $ChunkTrimming = @()
                $ChunkTrimming2 = @()
                $ChunkObject = @()
                $GetNextChunk = @()
            
                $GetNextChunk = Get-Content $HDD_Alert_Combined | Select-Object -First 8

                $ChunkTrimming = $GetNextChunk.split(": ")
                $ChunkTrimming2 = $ChunkTrimming.trim() -ne ""
                $ChunkTrimmed = $ChunkTrimming2 | Where-Object { (-not($_ -match 'HDD_*')) }

                # Build CSV and append
                $ChunkObject = New-Object PSObject
                    $ChunkObject | Add-Member -Type NoteProperty -Name HDD_Target –Value $ChunkTrimmed[0]
                    $ChunkObject | Add-Member -type NoteProperty -name HDD_DiskName -Value $ChunkTrimmed[1]
                    $ChunkObject | Add-Member -type NoteProperty -name HDD_DiskFree -Value $ChunkTrimmed[2]
                    $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskUsed –Value $ChunkTrimmed[3]
                    $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskMax –Value $ChunkTrimmed[4]
                    $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskFreePct –Value $ChunkTrimmed[5]
                    $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskThreshold –Value $ChunkTrimmed[6]
                    $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskAlertFlag –Value $ChunkTrimmed[7]
            
                cd $ReportingPath

                $HDDAlertPath = "Alert_HDD_Report_" + "$DateStamp" + ".CSV"
                $HDDAlertConvert = ConvertTo-CSV -InputObject $ChunkObject -NoTypeInformation
                Export-CSV "$HDDAlertPath" -InputObject $ChunkObject -Append -NoTypeInformation

                # Remove lines from HDDSpecCombined.txt and start Chunk over
                cd $TempPath
                Get-content $HDD_Alert_Combined | select-object -skip 8 > "NextAlertChunkList.txt"
                Remove-Item $HDD_Alert_Combined -Force
                Rename-Item NextAlertChunkList.txt $HDD_Alert_Combined
            
            }
        
            While ($GetNextChunk -ne $NULL)
        }


        # HDD Spec Consolidation and Report
        cd $TempPath
        $TempReportPull = @()

        # Remove blank lines
        $TempReportPull = Get-ChildItem | Where {$_.Name -LIKE "HDDSpecs_*"}
        Get-Content $TempReportPull | where {$_.trim() -ne "" } > "HDDSpecCombined.txt"
        $HDDSpecCombined = "HDDSpecCombined.txt"

        Do { 
            # Trim away unneeded data from import
            Clear-Variable -Name ChunkTrimmed,ChunkTrimming,ChunkTrimming2,ChunkObject,GetNextChunk
            $ChunkTrimmed = @()
            $ChunkTrimming = @()
            $ChunkTrimming2 = @()
            $ChunkObject = @()
            $GetNextChunk = @()
            
            $GetNextChunk = Get-Content $HDDSpecCombined | Select-Object -First 8

            $ChunkTrimming = $GetNextChunk.split(": ")
            $ChunkTrimming2 = $ChunkTrimming.trim() -ne ""
            $ChunkTrimmed = $ChunkTrimming2 | Where-Object { (-not($_ -match 'HDD_*')) }

            # Build CSV and append
            $ChunkObject = New-Object PSObject
                $ChunkObject | Add-Member -MemberType NoteProperty -Name HDD_Target –Value $ChunkTrimmed[0]
                $ChunkObject | Add-Member -MemberType NoteProperty -name HDD_DiskName -Value $ChunkTrimmed[1]
                $ChunkObject | Add-Member -MemberType NoteProperty -name HDD_DiskFree -Value $ChunkTrimmed[2]
                $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskUsed –Value $ChunkTrimmed[3]
                $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskMax –Value $ChunkTrimmed[4]
                $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskFreePct –Value $ChunkTrimmed[5]
                $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskThreshold –Value $ChunkTrimmed[6]
                $ChunkObject | Add-Member –MemberType NoteProperty -Name HDD_DiskAlertFlag –Value $ChunkTrimmed[7]
            
            cd $ReportingPath

            $HDDChunkPath = "Host_HDD_Report_" + "$DateStamp" + ".CSV"
            $HDDConvert = ConvertTo-CSV -InputObject $ChunkObject -NoTypeInformation
            Export-CSV "$HDDChunkPath" -InputObject $ChunkObject -Append -NoTypeInformation

            # Remove lines from HDDSpecCombined.txt and start Chunk over
            cd $TempPath
            Get-content $HDDSpecCombined | select-object -skip 8 > "NextChunkList.txt"
            Remove-Item $HDDSpecCombined -Force
            Rename-Item NextChunkList.txt $HDDSpecCombined
            
        }
        
        While ($GetNextChunk -ne $NULL)
        
        $error.Clear()
        cd $ReportingPath
        Remove-Item Temp -Force -Recurse
        
        clear

        # Output ############################################################################################
        Write-Host "~~~~~~~~~~~~~~~~~~" -ForegroundColor Magenta
        Write-Host "~~~ Host Audit ~~~" -ForegroundColor Magenta
        Write-Host "~~~~~~~~~~~~~~~~~~" -ForegroundColor Magenta
        Write-Host
        Write-Host "The following computers were scanned:" -ForegroundColor Cyan
        Write-Host "=====================================" -ForegroundColor Cyan
        $HostNameStamp
        Write-Host 
        Write-Host 
        Write-Host "Errors Reported" -ForegroundColor Cyan
        Write-Host "===============" -ForegroundColor Cyan
        Write-Host
       
        # Define Alerts #######################################################
        $EmailNeeded = "0"
        
        #if RAM error...
        Write-Host "Problems with RAM" -ForegroundColor Yellow
        if ($RAM_Alerting -eq $TRUE) {
            Write-Host "Error Found:" -ForegroundColor Red
            $EmailNeeded = "1"
            $RAM_Alert = @()
            $RAM_Alert = Import-Csv ".\Alert_RAM_Report_$DateStamp.CSV"
            $RAM_Alert | ft * -AutoSize
        }
        else {
            Write-Host "No RAM problems reported!" -ForegroundColor Green
            Write-Host
        }

        #if CPU error...
        Write-Host "Problems with CPU" -ForegroundColor Yellow
        if ($CPU_Alerting -eq $TRUE) {
            Write-Host "Error Found:" -ForegroundColor Red
            $EmailNeeded = "1"
            $CPU_Alert = @()
            $CPU_Alert = Import-Csv ".\Alert_CPU_Report_$DateStamp.CSV"
            $CPU_Alert | ft * -AutoSize
        }
        else {
            Write-Host "No CPU problems reported!" -ForegroundColor Green
            Write-Host
        }


        #if HDD error...
        Write-Host "Problems with Disk Space" -ForegroundColor Yellow
        if ($HDD_Alerting -eq $TRUE) {
            Write-Host "Error Found:" -ForegroundColor Red
            $EmailNeeded = "1"
            $HDD_Alert = @()
            $HDD_Alert = Import-Csv ".\Alert_HDD_Report_$DateStamp.CSV"
            $HDD_Alert | ft * -AutoSize
        }
        else {
            Write-Host "No HDD problems reported!" -ForegroundColor Green
            Write-Host
        }


        
        # Email for errors ##############################################################
        if ($EmailNeeded -eq "1") { 
            
            $RAM_Path = "$ReportingPath\Alert_RAM_Report_$DateStamp.CSV"
            $CPU_Path = "$ReportingPath\Alert_CPU_Report_$DateStamp.CSV"
            $HDD_Path = "$ReportingPath\Alert_HDD_Report_$DateStamp.CSV"
            
            #Email Body
            if (Test-Path $RAM_Path) {
                $CSVMAILBODY1 = (Import-CSV $RAM_Path | Out-String)
            }
            else {
                $CSVMAILBODY1 = "No RAM Issues Reported"
            }

            if (Test-Path $CPU_Path) {
                $CSVMAILBODY2 = (Import-CSV $CPU_Path | Out-String)
            }
            else {
                $CSVMAILBODY2 = "No CPU Issues Reported"
            }

            if (Test-Path $HDD_Path) {
                $CSVMAILBODY3 = (Import-CSV $HDD_Path | Out-String)
            }
            else {
                $CSVMAILBODY3 = "No HDD Issues Reported"
            }

$CVSMailBodyFull = "This is an automated message generated by a PowerShell script running from $env:COMPUTERNAME. 

It will include data from the following:
RAM Errors:
$CSVMAILBODY1

CPU Errors:
$CSVMAILBODY2

HDD Errors:
$CSVMAILBODY3
"
            
            Send-MailMessage -To $To -Subject "Host Error Report" -Body $CVSMailBodyFull -SmtpServer $SMTPSERVER -From $From

        }

    }

    #Run Spec Scan
    HostSpecs
    
}
