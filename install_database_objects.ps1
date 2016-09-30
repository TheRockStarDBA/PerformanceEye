#####
# PRODUCT:    PerformanceEye  https://github.com/amorelli005/PerformanceEye
# PROCEDURE:	install_database_objects.ps1
#
# AUTHOR: Aaron Morelli
#	email@TBD.com
#	@sqlcrossjoin
#	sqlcrossjoin.wordpress.com					
#
#	PURPOSE: Install the PerformanceEye database, its schema and T-SQL objects, and SQL Agent jobs
#  on a SQL Server instance.
#
#	OUTSTANDING ISSUES: None at this time.
#
#   CHANGE LOG:	
#		2016-09-30	Aaron Morelli		Final code run-through and commenting
#
#	MIT License
#
#	Copyright (c) 2016 Aaron Morelli
#
#	Permission is hereby granted, free of charge, to any person obtaining a copy
#	of this software and associated documentation files (the "Software"), to deal
#	in the Software without restriction, including without limitation the rights
#	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#	copies of the Software, and to permit persons to whom the Software is
#	furnished to do so, subject to the following conditions:
#
#	The above copyright notice and this permission notice shall be included in all
#	copies or substantial portions of the Software.
#
#	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#	SOFTWARE.

# To Execute
# ------------------------
# <this should not be called directly. Call the parent PerformanceEye_Installer.ps1 
# script instead
#####

param ( 
[Parameter(Mandatory=$true)][string]$Server, 
[Parameter(Mandatory=$true)][string]$Database,
[Parameter(Mandatory=$true)][string]$HoursToKeep,
[Parameter(Mandatory=$true)][string]$curScriptLocation
) 

Import-Module SqlPs



$core_parent = $curScriptLocation + "core\"
$core_functions = $core_parent + "functions"
$core_procedures = $core_parent + "procedures"
$core_schemas = $core_parent + "schemas"
$core_tables = $core_parent + "tables"
$core_views = $core_parent + "views"

$autowho_parent = $curScriptLocation + "autowho\"
$autowho_functions = $autowho_parent + "functions"
$autowho_procedures = $autowho_parent + "procedures"
$autowho_schemas = $autowho_parent + "schemas"
$autowho_tables = $autowho_parent + "tables"
$autowho_types = $autowho_parent + "types"
$autowho_views = $autowho_parent + "views"

$servereye_parent = $curScriptLocation + "servereye\"
$servereye_functions = $servereye_parent + "functions"
$servereye_procedures = $servereye_parent + "procedures"
$servereye_schemas = $servereye_parent + "schemas"
$servereye_tables = $servereye_parent + "tables"
$servereye_views = $servereye_parent + "views"

$job_core = $curScriptLocation + "jobs\CTCoreTraceMaster.sql"
$job_autowho = $curScriptLocation + "jobs\AutoWhoTrace.sql"

$masterprocs_parent = $curScriptLocation + "masterprocs\"



# check whether DB exists. If not, create it.
try {
	$MyVariableArray = "DBName = $Database"
	
	invoke-sqlcmd -inputfile ".\CreateCTDatabase.sql" -serverinstance $Server -database master -Variable $MyVariableArray -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished if-not-exists-create check for $Database" -foregroundcolor cyan -backgroundcolor black
}
catch [system.exception] {
	write-host "Error occurred in CreateCTDatabase.sql: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}


# we take our __TEMPLATE versions of the master procs and create versions with $Database substituted for MoTools
$masterproc_LR = $masterprocs_parent + "sp_LongRequests__TEMPLATE.sql"
$masterproc_LR_Replace = $masterprocs_parent + "sp_LongRequests__" + $Database + ".sql"

if (Test-Path $masterproc_LR_Replace) {
	Remove-Item $masterproc_LR_Replace
}

(Get-Content $masterproc_LR) | Foreach-Object { $_ -replace 'MoTools', $Database } | Set-Content $masterproc_LR_Replace

try {
	invoke-sqlcmd -inputfile $masterproc_LR_Replace -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished creating sp_LongRequests" -foregroundcolor cyan -backgroundcolor black

}
catch [system.exception] {
	write-host "Error occurred while creating sp_LongRequests: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}


$masterproc_SS = $masterprocs_parent + "sp_SessionSummary__TEMPLATE.sql"
$masterproc_SS_Replace = $masterprocs_parent + "sp_SessionSummary__" + $Database + ".sql"

if (Test-Path $masterproc_SS_Replace) {
	Remove-Item $masterproc_SS_Replace
}

(Get-Content $masterproc_SS) | Foreach-Object { $_ -replace 'MoTools', $Database } | Set-Content $masterproc_SS_Replace

try {
	invoke-sqlcmd -inputfile $masterproc_SS_Replace -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished creating sp_SessionSummary" -foregroundcolor cyan -backgroundcolor black

}
catch [system.exception] {
	write-host "Error occurred while creating sp_SessionSummary: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}

$masterproc_AS = $masterprocs_parent + "sp_AutoSummary__TEMPLATE.sql"
$masterproc_AS_Replace = $masterprocs_parent + "sp_AutoSummary__" + $Database + ".sql"

if (Test-Path $masterproc_AS_Replace) {
	Remove-Item $masterproc_AS_Replace
}

(Get-Content $masterproc_AS) | Foreach-Object { $_ -replace 'MoTools', $Database } | Set-Content $masterproc_AS_Replace

try {
	invoke-sqlcmd -inputfile $masterproc_AS_Replace -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished creating sp_AutoSummary" -foregroundcolor cyan -backgroundcolor black
}
catch [system.exception] {
	write-host "Error occurred while creating sp_AutoSummary: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}

$masterproc_SV = $masterprocs_parent + "sp_SessionViewer__TEMPLATE.sql"
$masterproc_SV_Replace = $masterprocs_parent + "sp_SessionViewer__" + $Database + ".sql"

if (Test-Path $masterproc_SV_Replace) {
	Remove-Item $masterproc_SV_Replace
}

(Get-Content $masterproc_SV) | Foreach-Object { $_ -replace 'MoTools', $Database } | Set-Content $masterproc_SV_Replace

try {
	invoke-sqlcmd -inputfile $masterproc_SV_Replace -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished creating sp_SessionViewer" -foregroundcolor cyan -backgroundcolor black

}
catch [system.exception] {
	write-host "Error occurred while creating sp_SessionViewer: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}


# Schemas  
try {
	(dir $core_schemas) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation

		}
        
	(dir $autowho_schemas) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
        
    #TODO: add code to create ServerEye schema
}
catch [system.exception] {
	write-host "Error occurred when creating schemas: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of Schema block



# Core tables
try {
	(dir $core_tables) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation

		}
}
catch [system.exception] {
	write-host "Error occurred when creating core tables: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of Core Tables block


# Core Views
try {
	(dir $core_views) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating core views: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of Core views


# Core functions
try {
	(dir $core_functions) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating core functions: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of Core functions


# Core procedures
try {
	(dir $core_procedures) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating core procedures: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of Core procedures



# AutoWho types
try {
	(dir $autowho_types) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating AutoWho types: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of AutoWho types



# AutoWho tables
try {
	(dir $autowho_tables) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating AutoWho tables: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of AutoWho tables block



# AutoWho Views
try {
	(dir $autowho_views) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating AutoWho views: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of AutoWho views



# AutoWho functions
try {
	(dir $autowho_functions) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating AutoWho functions: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of AutoWho functions


# AutoWho procedures
try {
	(dir $autowho_procedures) |  
		ForEach-Object {  
			$curScript = $_.FullName
			$curFileName = $_.Name

			invoke-sqlcmd -inputfile $curScript -serverinstance $Server -database $Database -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
			#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
			CD $curScriptLocation
		}
}
catch [system.exception] {
	write-host "Error occurred when creating AutoWho procedures: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
}  # end of AutoWho procedures


# AutoWho config
try {
	$MyVariableArray = "HoursToKeep = $HoursToKeep"
	
	invoke-sqlcmd -inputfile ".\AutoWhoConfig.sql" -serverinstance $Server -database $Database -Variable $MyVariableArray -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished AutoWho configuration" -foregroundcolor cyan -backgroundcolor black
}
catch [system.exception] {
	write-host "Error occurred in AutoWhoConfig.sql: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
} # AutoWho config


# Trace Master job
try {
	$MyVariableArray = "DBName = $Database"
	
	invoke-sqlcmd -inputfile $job_core -serverinstance $Server -database msdb -Variable $MyVariableArray -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished creating Trace Master job" -foregroundcolor cyan -backgroundcolor black
}
catch [system.exception] {
	write-host "Error occurred when creating Trace Master job: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
} # Trace Master job


# AutoWho job
try {
	$MyVariableArray = "DBName = $Database"
	
	invoke-sqlcmd -inputfile $job_autowho -serverinstance $Server -database msdb -Variable $MyVariableArray -QueryTimeout 65534 -AbortOnError -Verbose -outputsqlerrors $true
	#In Windows 2012 R2, we are ending up in the SQLSERVER:\ prompt, when really we want to be in the file system provider. Doing a simple "CD" command gets us back there
	CD $curScriptLocation

	Write-Host "Finished creating AutoWho job" -foregroundcolor cyan -backgroundcolor black
}
catch [system.exception] {
	write-host "Error occurred when creating AutoWho job: " -foregroundcolor red -backgroundcolor black
	Write-host "$_" -foregroundcolor red -backgroundcolor black
	write-host "Exiting..." -foregroundcolor red -backgroundcolor black
	break
} # AutoWho job
