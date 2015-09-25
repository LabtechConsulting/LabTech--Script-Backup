<#
.SYNOPSIS
    Backs up all LabTech scripts.

.DESCRIPTION
    This script will export all LabTech sctipts in xml format to a specified destination.
    Requires the MySQL .NET connector.

.LINK
        http://www.labtechconsulting.com
        https://dev.mysql.com/downloads/connector/net/6.9.html
        
.OUTPUTS
    Default values -
    Log file stored in: $($env:windir)\LTScv\Logs\LT-ScriptExport.log
    Scripts exported to: $($env:windir)\Program Files(x86)\LabTech\Backup\Scripts
    Credentials file: $PSScriptRoot

.NOTES
    Version:        1.0
    Author:         Chris Taylor
    Website:        www.labtechconsulting.com
    Creation Date:  9/11/2015
    Purpose/Change: Initial script development

    Version:        1.1
    Author:         Chris Taylor
    Website:        www.labtechconsulting.com
    Creation Date:  9/23/2015
    Purpose/Change: Added error catching
#>

#Requires -Version 3.0 
 
#region-[Declarations]----------------------------------------------------------
    
    $ScriptVersion = "1.1"
    
    $ErrorActionPreference = "Stop"
    
    $Date = Get-Date
    
    #Get/Save config info
    if($(Test-Path $PSScriptRoot\LT-ScriptExport-Config.xml) -eq $false) {
        #Config file template
        $Config = [xml]@'
<Settings>
	<LogPath></LogPath>
	<BackupRoot></BackupRoot>
	<MySQLDatabase></MySQLDatabase>
	<MySQLHost></MySQLHost>
	<CredPath></CredPath>
    <LastExport>0</LastExport>
</Settings>
'@
        try {
            #Create config file
            $Config.Settings.LogPath = "$(Read-Host "Path of log file ($($env:windir)\LTSvc\Logs)")"
            if ($Config.Settings.LogPath -eq '') {$Config.Settings.LogPath = "$($env:windir)\LTSvc\Logs"}
            $Config.Settings.BackupRoot = "$(Read-Host "Path of exported scripts (${env:ProgramFiles(x86)}\LabTech\Backup\Scripts)")"
            if ($Config.Settings.BackupRoot -eq '') {$Config.Settings.BackupRoot = "${env:ProgramFiles(x86)}\LabTech\Backup\Scripts"}
            $Config.Settings.MySQLDatabase = "$(Read-Host "Name of LabTech database (labtech)")"
            if ($Config.Settings.MySQLDatabase -eq '') {$Config.Settings.MySQLDatabase = "labtech"}
            $Config.Settings.MySQLHost = "$(Read-Host "FQDN of LabTechServer (localhost)")"
            if ($Config.Settings.MySQLHost -eq '') {$Config.Settings.MySQLHost = "localhost"}
            $Config.Settings.CredPath = "$(Read-Host "Path of credentials ($PSScriptRoot)")"
            if ($Config.Settings.CredPath -eq '') {$Config.Settings.CredPath = "$PSScriptRoot"}
            $Config.Save("$PSScriptRoot\LT-ScriptExport-Config.xml")
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Log-Error -LogPath $FullLogPath  -ErrorDesc "Error durring config creation: $FailedItem, $ErrorMessage `n$Error[0]" -ExitGracefully $True
        }
    }
    Else {
    [xml]$Config = Get-Content "$PSScriptRoot\LT-ScriptExport-Config.xml"
    }

    #Location to credentials file
    $CredPath = $Config.Settings.CredPath
    
    #Get/Save user/password info
    if ($(Test-Path $CredPath) -eq $false) {New-Item -ItemType Directory -Force -Path $CredPath | Out-Null}
    if($(Test-Path $CredPath\LTDBCredentials.xml) -eq $false){Get-Credential -Message "Please provide the credentials to the LabTech MySQL database." | Export-Clixml $CredPath\LTDBCredentials.xml -Force}
    
    #Log File Info
    $LogName = "LT-ScriptExport.log"
    $LogPath = ($Config.Settings.LogPath)
    $FullLogPath = $LogPath + "\" + $LogName


    #Location to the backp repository
    $BackupRoot = $Config.Settings.BackupRoot

    #MySQL connection info
    $MySQLDatabase = $Config.Settings.MySQLDatabase
    $MySQLHost = $Config.Settings.MySQLHost
    $MySQLAdminPassword = (IMPORT-CLIXML "$CredPath\LTDBCredentials.xml").GetNetworkCredential().Password
    $MySQLAdminUserName = (IMPORT-CLIXML "$CredPath\LTDBCredentials.xml").GetNetworkCredential().UserName

#endregion
 
#region-[Functions]------------------------------------------------------------

Function Log-Start{
  <#
  .SYNOPSIS
    Creates log file

  .DESCRIPTION
    Creates log file with path and name that is passed. Checks if log file exists, and if it does deletes it and creates a new one.
    Once created, writes initial logging data

  .PARAMETER LogPath
    Mandatory. Path of where log is to be created. Example: C:\Windows\Temp

  .PARAMETER LogName
    Mandatory. Name of log file to be created. Example: Test_Script.log
      
  .PARAMETER ScriptVersion
    Mandatory. Version of the running script which will be written in the log. Example: 1.5

  .INPUTS
    Parameters above

  .OUTPUTS
    Log file created

  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development

    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support

    Version:        1.2
    Author:         Chris Taylor
    Creation Date:  7/17/2015
    Purpose/Change: Added directory creation if not present.
                    Added Append option
                    


  .EXAMPLE
    Log-Start -LogPath "C:\Windows\Temp" -LogName "Test_Script.log" -ScriptVersion "1.5"
  #>
    
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$LogPath, [Parameter(Mandatory=$true)][string]$LogName, [Parameter(Mandatory=$true)][string]$ScriptVersion, [Parameter(Mandatory=$false)][switch]$Append)
  
  Process{
    $FullLogPath = $LogPath + "\" + $LogName
    #Check if file exists and delete if it does
    If((Test-Path -Path $FullLogPath) -and $Append -ne $true){
      Remove-Item -Path $FullLogPath -Force
    }

    #Check if folder exists if not create    
    If((Test-Path -PathType Container -Path $LogPath) -eq $False){
      New-Item -ItemType Directory -Force -Path $LogPath
    }

    #Create file and start logging
    If($(Test-Path -Path $FullLogPath) -ne $true) {
        New-Item -Path $LogPath -Value $LogName -ItemType File
    }

    Add-Content -Path $FullLogPath -Value "***************************************************************************************************"
    Add-Content -Path $FullLogPath -Value "Started processing at [$([DateTime]::Now)]."
    Add-Content -Path $FullLogPath -Value "***************************************************************************************************"
    Add-Content -Path $FullLogPath -Value ""
    Add-Content -Path $FullLogPath -Value "Running script version [$ScriptVersion]."
    Add-Content -Path $FullLogPath -Value ""
    Add-Content -Path $FullLogPath -Value "***************************************************************************************************"
    Add-Content -Path $FullLogPath -Value ""
  
    #Write to screen for debug mode
    Write-Debug "***************************************************************************************************"
    Write-Debug "Started processing at [$([DateTime]::Now)]."
    Write-Debug "***************************************************************************************************"
    Write-Debug ""
    Write-Debug "Running script version [$ScriptVersion]."
    Write-Debug ""
    Write-Debug "***************************************************************************************************"
    Write-Debug ""
  }
}
 
Function Log-Write{
  <#
  .SYNOPSIS
    Writes to a log file

  .DESCRIPTION
    Appends a new line to the end of the specified log file
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Test_Script.log
  
  .PARAMETER LineValue
    Mandatory. The string that you want to write to the log
      
  .INPUTS
    Parameters above

  .OUTPUTS
    None

  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development
  
    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support

  .EXAMPLE
    Log-Write -FullLogPath "C:\Windows\Temp\Test_Script.log" -LineValue "This is a new line which I am appending to the end of the log file."
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$FullLogPath, [Parameter(Mandatory=$true)][string]$LineValue)
  
  Process{
    Add-Content -Path $FullLogPath -Value $LineValue
    
    Write-Output $LineValue

    #Write to screen for debug mode
    Write-Debug $LineValue
  }
}
 
Function Log-Error{
  <#
  .SYNOPSIS
    Writes an error to a log file

  .DESCRIPTION
    Writes the passed error to a new line at the end of the specified log file
  
  .PARAMETER FullLogPath
    Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Test_Script.log
  
  .PARAMETER ErrorDesc
    Mandatory. The description of the error you want to pass (use $_.Exception)
  
  .PARAMETER ExitGracefully
    Mandatory. Boolean. If set to True, runs Log-Finish and then exits script

  .INPUTS
    Parameters above

  .OUTPUTS
    None

  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development
    
    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support. Added -ExitGracefully parameter functionality
                    
  .EXAMPLE
    Log-Error -FullLogPath "C:\Windows\Temp\Test_Script.log" -ErrorDesc $_.Exception -ExitGracefully $True
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$FullLogPath, [Parameter(Mandatory=$true)][string]$ErrorDesc, [Parameter(Mandatory=$true)][boolean]$ExitGracefully)
  
  Process{
    Add-Content -Path $FullLogPath -Value "Error: An error has occurred [$ErrorDesc]."
  
    Write-Error $ErrorDesc

    #Write to screen for debug mode
    Write-Debug "Error: An error has occurred [$ErrorDesc]."
    
    #If $ExitGracefully = True then run Log-Finish and exit script
    If ($ExitGracefully -eq $True){
      Log-Finish -FullLogPath $FullLogPath
      Breaåk
    }
  }
}
 
Function Log-Finish{
  <#
  .SYNOPSIS
    Write closing logging data & exit

  .DESCRIPTION
    Writes finishing logging data to specified log and then exits the calling script
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write finishing data to. Example: C:\Windows\Temp\Test_Script.log

  .PARAMETER NoExit
    Optional. If this is set to True, then the function will not exit the calling script, so that further execution can occur
  
  .PARAMETER Limit
    Optional. Sets the max linecount of the script.
  
  .INPUTS
    Parameters above

  .OUTPUTS
    None

  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development
    
    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support
  
    Version:        1.2
    Author:         Luca Sturlese
    Creation Date:  01/08/12
    Purpose/Change: Added option to not exit calling script if required (via optional parameter)

    Version:        1.3
    Author:         Chris Taylor
    Creation Date:  7/17/2015
    Purpose/Change: Added log line count limit.
    
  .EXAMPLE
    Log-Finish -FullLogPath "C:\Windows\Temp\Test_Script.log"

.EXAMPLE
    Log-Finish -FullLogPath "C:\Windows\Temp\Test_Script.log" -NoExit $True
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$FullLogPath, [Parameter(Mandatory=$false)][string]$NoExit, [Parameter(Mandatory=$false)][int]$Limit )
  
  Process{
    Add-Content -Path $FullLogPath -Value ""
    Add-Content -Path $FullLogPath -Value "***************************************************************************************************"
    Add-Content -Path $FullLogPath -Value "Finished processing at [$([DateTime]::Now)]."
    Add-Content -Path $FullLogPath -Value "***************************************************************************************************"
  
    #Write to screen for debug mode
    Write-Debug ""
    Write-Debug "***************************************************************************************************"
    Write-Debug "Finished processing at [$([DateTime]::Now)]."
    Write-Debug "***************************************************************************************************"
  
    if ($Limit){
        #Limit Log file to 50000 lines
        (Get-Content $FullLogPath -tail $Limit -readcount 0) | Set-Content $FullLogPath -Force -Encoding Unicode
    }
    #Exit calling script if NoExit has not been specified or is set to False
    If(!($NoExit) -or ($NoExit -eq $False)){
      Exit
    }    
  }
} 

Function Get-LTData {
    <#
    .SYNOPSIS
        Executes a MySQL query aginst the LabTech Databse.

    .DESCRIPTION
        This comandlet will execute a MySQL query aginst the LabTech database.
        Requires the MySQL .NET connector.
        Original script by Dan Rose

    .LINK
        https://dev.mysql.com/downloads/connector/net/6.9.html
        https://www.cogmotive.com/blog/powershell/querying-mysql-from-powershell
        http://www.labtechconsulting.com

    .PARAMETER Query
        Input your MySQL query in double quotes.

    .INPUTS
        Pipeline

    .NOTES
        Version:        1.0
        Author:         Chris Taylor
        Website:        www.labtechconsulting.com
        Creation Date:  9/11/2015
        Purpose/Change: Initial script development
  
    .EXAMPLE
        Get-LTData "SELECT ScriptID FROM lt_scripts"
        $Query | Get-LTData
    #>

    Param(
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true)]
        [string]$Query
    )

    Begin {
        $ConnectionString = "server=" + $MySQLHost + ";port=3306;uid=" + $MySQLAdminUserName + ";pwd=" + $MySQLAdminPassword + ";database="+$MySQLDatabase
    }

    Process {
        Try {
          [void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
          $Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
          $Connection.ConnectionString = $ConnectionString
          $Connection.Open()

          $Command = New-Object MySql.Data.MySqlClient.MySqlCommand($Query, $Connection)
          $DataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($Command)
          $DataSet = New-Object System.Data.DataSet
          $RecordCount = $dataAdapter.Fill($dataSet, "data")
          $DataSet.Tables[0]
        }

        Catch {
          Log-Error -FullLogPath $FullLogPath  -ErrorDesc "Unable to run query : $query `n$Error[0]" -ExitGracefully $True
        }
    }

    End {
      $Connection.Close()
    }
}

Function Export-LTScript {
    <#
    .SYNOPSIS
        Exports a LabTech script as an xml file.

    .DESCRIPTION
        This commandlet will execute a MySQL query aginst the LabTech database.
        Requires Get-LTData
        
    .LINK
        http://www.labtechconsulting.com

    .PARAMETER Query
        Input your MySQL query in double quotes.
    
    .PARAMETER FilePath
        File path of exported script.

    .NOTES
        Version:        1.0
        Author:         Chris Taylor
        Website:        www.labtechconsulting.com
        Creation Date:  9/11/2015
        Purpose/Change: Initial script development
  
    .EXAMPLE
        Get-LTData "SELECT ScriptID FROM lt_scripts" -FilePath C:\Windows\Temp
    #>

    [CmdletBinding()]
        Param(
        [Parameter(Mandatory=$True,Position=1)]
        [string]$ScriptID
    )

    #LabTech XML template
    $ExportTemplate = [xml] @"
<LabTech_Expansion
	Version="100.332"
	Name="LabTech Script Expansion"
	Type="PackedScript">
	<PackedScript>
		<NewDataSet>
			<Table>
				<ScriptId></ScriptId>
				<FolderId></FolderId>
				<ScriptName></ScriptName>
				<ScriptNotes></ScriptNotes>
				<Permission></Permission>
				<EditPermission></EditPermission>
				<ComputerScript></ComputerScript>
				<LocationScript></LocationScript>
				<MaintenanceScript></MaintenanceScript>
				<FunctionScript></FunctionScript>
				<LicenseData></LicenseData>
				<ScriptData></ScriptData>
				<ScriptVersion></ScriptVersion>
				<ScriptGuid></ScriptGuid>
				<ScriptFlags></ScriptFlags>
				<Parameters></Parameters>
			</Table>
		</NewDataSet>
		<ScriptFolder>
			<NewDataSet>
				<Table>
					<FolderID></FolderID>
					<ParentID></ParentID>
					<Name></Name>
					<GUID></GUID>
				</Table>
			</NewDataSet>
		</ScriptFolder>
	</PackedScript>
</LabTech_Expansion>
"@

    #Query MySQL for script data.
    $ScriptXML = Get-LTData -query "SELECT * FROM lt_scripts WHERE ScriptID=$ScriptID"
    $ScriptData = Get-LTData -query "SELECT CONVERT(ScriptData USING utf8) AS Data FROM lt_scripts WHERE ScriptID=$ScriptID"
    $ScriptLicense = Get-LTData -query "SELECT CONVERT(LicenseData USING utf8) AS License FROM lt_scripts WHERE ScriptID=$ScriptID"

    #Save script data to the template.
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ScriptId = "$($ScriptXML.ScriptId)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.FolderId = "$($ScriptXML.FolderId)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ScriptName = "$($ScriptXML.ScriptName)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ScriptNotes = "$($ScriptXML.ScriptNotes)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.Permission = "$($ScriptXML.Permission)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.EditPermission = "$($ScriptXML.EditPermission)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ComputerScript = "$($ScriptXML.ComputerScript)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.LocationScript = "$($ScriptXML.LocationScript)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.MaintenanceScript = "$($ScriptXML.MaintenanceScript)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.FunctionScript = "$($ScriptXML.FunctionScript)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.LicenseData = "$($ScriptLicense.License)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ScriptData = "$($ScriptData.Data)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ScriptVersion = "$($ScriptXML.ScriptVersion)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ScriptGuid = "$($ScriptXML.ScriptGuid)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.ScriptFlags = "$($ScriptXML.ScriptFlags)"
    $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.Parameters = "$($ScriptXML.Parameters)"
    
    #Format Script Name
    #Remove special characters
    $FileName = $($ScriptXML.ScriptName).Replace('*','')
    $FileName = $FileName.Replace('/','-')
    $FileName = $FileName.Replace('<','')
    $FileName = $FileName.Replace('>','')
    $FileName = $FileName.Replace(':','')
    $FileName = $FileName.Replace('"','')
    $FileName = $FileName.Replace('\','-')
    $FileName = $FileName.Replace('|','')
    $FileName = $FileName.Replace('?','')
    #Add last modification date
    $FileName = $($FileName) + '--' + $($ScriptXML.Last_Date.ToString("yyyy-MM-dd--HH-mm-ss"))
    #Add last user to modify
    $FileName = $($FileName) + '--' + $($ScriptXML.Last_User.Substring(0, $ScriptXML.Last_User.IndexOf('@')))
    

    #Check folder information

    #Check if script is at root and not in a folder
    If ($($ScriptXML.FolderId) -eq 0) {
        try {
            #Delete folder information from template
            $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.RemoveAttribute()
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Log-Error -FullLogPath $FullLogPath  -ErrorDesc "Unable to remove folder data from XML: $FailedItem, $ErrorMessage `n$Error[0]" -ExitGracefully $True
        }
    }
    Else {    
        #Query MySQL for folder data.
        $FolderData = Get-LTData -query "SELECT * FROM `scriptfolders` WHERE FolderID=$($ScriptXML.FolderId)"
        
        #Check if folder is no longer present. 
        if ($FolderData -eq $null) {
            Log-Write -FullLogPath $FullLogPath -LineValue "ScritID $($ScriptXML.ScriptId) references folder $($ScriptXML.FolderId), this folder is no longer present. Setting to root folder."
            Log-Write -FullLogPath $FullLogPath -LineValue "It is recomended that you move this script to a folder."
            
            #Set to FolderID 0
            $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.FolderId = "0"
            $ScriptXML.FolderID = 0
            
            try {            
                #Delete folder information from template
                $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.RemoveAttribute()
            }
            Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Log-Error -FullLogPath $FullLogPath  -ErrorDesc "Unable to remove folder data from XML: $FailedItem, $ErrorMessage `n$Error[0]" -ExitGracefully $True
            }
        }
        Else {
            #Format the folder name.
            #Remove special characters
            $FolderName = $($FolderData.Name).Replace('*','')
            $FolderName = $FolderName.Replace('/','-')
            $FolderName = $FolderName.Replace('<','')
            $FolderName = $FolderName.Replace('>','')
            $FolderName = $FolderName.Replace(':','')
            $FolderName = $FolderName.Replace('"','')
            $FolderName = $FolderName.Replace('\','-')
            $FolderName = $FolderName.Replace('|','')
            $FolderName = $FolderName.Replace('?','')
            
            # Save folder data to the template.
            $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.NewDataSet.Table.FolderID = "$($FolderData.FolderID)"
            $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.NewDataSet.Table.ParentID = "$($FolderData.ParentID)"
            $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.NewDataSet.Table.Name = "$FolderName"
            $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.NewDataSet.Table.GUID = "$($FolderData.GUID)"
        }
    }

    #Create Folder Structure. Check for parent folder 
    If ($($FolderData.ParentId) -eq 0) {
        try {
            #Create folder
            New-Item -ItemType Directory -Force -Path $BackupRoot\$($FolderData.Name) | Out-Null
        
            #Save XML
            $ExportTemplate.Save("$BackupRoot\$FolderName\$($FileName).xml")
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Log-Error -FullLogPath $FullLogPath  -ErrorDesc "Unable to save script: $FailedItem, $ErrorMessage `n$Error[0]" -ExitGracefully $True
        }
    }
    Else {
        #Query info for parent folder
        $ParentFolderName = $(Get-LTData "SELECT * FROM scriptfolders WHERE FolderID=$($FolderData.ParentID)").Name

        #Format parent folder name
        #Remove special characters
        $ParentFolderName = $ParentFolderName.Replace('*','')
        $ParentFolderName = $ParentFolderName.Replace('/','-')
        $ParentFolderName = $ParentFolderName.Replace('<','')
        $ParentFolderName = $ParentFolderName.Replace('>','')
        $ParentFolderName = $ParentFolderName.Replace(':','')
        $ParentFolderName = $ParentFolderName.Replace('"','')
        $ParentFolderName = $ParentFolderName.Replace('\','-')
        $ParentFolderName = $ParentFolderName.Replace('|','')
        $ParentFolderName = $ParentFolderName.Replace('?','')

        $FilePath = "$BackupRoot\$($ParentFolderName)\$($($FolderData.Name))"

        try {
            #Create folder
            New-Item -ItemType Directory -Force -Path $FilePath | Out-Null
        
            #Save XML
            $ExportTemplate.Save("$FilePath\$($FileName).xml")
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Log-Error -FullLogPath $FullLogPath  -ErrorDesc "Unable to save script: $FailedItem, $ErrorMessage `n$Error[0]" -ExitGracefully $True
        }
    }

}

#endregion

#region-[Execution]------------------------------------------------------------
    try {
    #Create log
    Log-Start -LogPath $LogPath -LogName $LogName -ScriptVersion $ScriptVersion -Append
   
    #Check backup directory
    if ((Test-Path $BackupRoot) -eq $false){New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null}
    }
    Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Log-Error -FullLogPath $FullLogPath  -ErrorDesc "Error durring log/backup directory creation: $FailedItem, $ErrorMessage `n$Error[0]" -ExitGracefully $True
        }
    
    Log-Write -FullLogPath $FullLogPath -LineValue "Getting list of all scripts."

    $ScriptIDs = @{}
    #Query list of all ScriptID's
    if ($($Config.Settings.LastExport) -eq 0) {
        $ScriptIDs = Get-LTData "SELECT ScriptID FROM lt_scripts"
    }
    else{
        $Query = $("SELECT ScriptID FROM lt_scripts WHERE Last_Date > " + "'" + $($Config.Settings.LastExport) +"'")
        $ScriptIDs = Get-LTData $Query   
    }
    
    Log-Write -FullLogPath $FullLogPath -LineValue "$(@($ScriptIDs).count) scripts to process."

    #Process each ScriptID
    $n = 0
    foreach ($ScriptID in $ScriptIDs) {
        #Progress bar
        $n++
        Write-Progress -Activity "Backing up LT scripts to $BackupRoot" -Status "Processing ScriptID $($ScriptID.ScriptID)" -PercentComplete  ($n / @($ScriptIDs).count*100)
        
        #Export current script
        Export-LTScript -ScriptID $($ScriptID.ScriptID)
    }

    Log-Write -FullLogPath $FullLogPath -LineValue "Export finished."
    
    try {
        $Config.Settings.LastExport = "$($Date.ToString("yyy-MM-dd HH:mm:ss"))"
        $Config.Save("$PSScriptRoot\LT-ScriptExport-Config.xml")
    }
    Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Log-Error -FullLogPath $FullLogPath  -ErrorDesc "Unable to update config with last export date: $FailedItem, $ErrorMessage `n$Error[0]" -ExitGracefully $True
        }

    Log-Finish -FullLogPath $FullLogPath

#endregion
