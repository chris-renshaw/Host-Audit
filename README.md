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
