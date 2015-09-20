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
#>

#Requires -Version 3.0 
 
#region-[Declarations]----------------------------------------------------------
    
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
    Else {
    [xml]$Config = Get-Content "$PSScriptRoot\LT-ScriptExport-Config.xml"
    }

    #Location to credentials file
    $CredPath = $Config.Settings.CredPath
    
    #Get/Save user/password info
    if ($(Test-Path $CredPath) -eq $false) {New-Item -ItemType Directory -Force -Path $CredPath | Out-Null}
    if($(Test-Path $CredPath\LTDBCredentials.xml) -eq $false){Get-Credential -Message "Please provide the credentials to the LabTech MySQL database." | Export-Clixml $CredPath\LTDBCredentials.xml -Force}
    
    #Transcript File Info
    $sTranscriptName = "LT-ScriptExport.log"
    $sTranscriptFile = ($Config.Settings.LogPath) + "\" + $sTranscriptName

    #Location to the backp repository
    $BackupRoot = $Config.Settings.BackupRoot

    #MySQL connection info
    $MySQLDatabase = $Config.Settings.MySQLDatabase
    $MySQLHost = $Config.Settings.MySQLHost
    $MySQLAdminPassword = (IMPORT-CLIXML $CredPath\LTDBCredentials.xml).GetNetworkCredential().Password
    $MySQLAdminUserName = (IMPORT-CLIXML $CredPath\LTDBCredentials.xml).GetNetworkCredential().UserName

#endregion
 
#region-[Functions]------------------------------------------------------------
 
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
          Write-Error "Unable to run query : $query `n$Error[0]"
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
        #Delete folder information from template
        $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.RemoveAttribute()
    }
    Else {    
        #Query MySQL for folder data.
        $FolderData = Get-LTData -query "SELECT * FROM `scriptfolders` WHERE FolderID=$($ScriptXML.FolderId)"
        
        #Check if folder is no longer present. 
        if ($FolderData -eq $null) {
            Write-Output "ScritID $($ScriptXML.ScriptId) references folder $($ScriptXML.FolderId), this folder is no longer present. Setting to root folder."
            Write-Output "It is recomended that you move this script to a folder."
            
            #Set to FolderID 0
            $ExportTemplate.LabTech_Expansion.PackedScript.NewDataSet.Table.FolderId = "0"
            $ScriptXML.FolderID = 0
                        
            #Delete folder information from template
            $ExportTemplate.LabTech_Expansion.PackedScript.ScriptFolder.RemoveAttribute()
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
        #Create folder
        New-Item -ItemType Directory -Force -Path $BackupRoot\$($FolderData.Name) | Out-Null
        
        #Save XML
        $ExportTemplate.Save("$BackupRoot\$FolderName\$($FileName).xml")
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

        $FilePath = "$BackupRoot\$($ParentFolderName)\$($ScriptFolderName)"

        #Create folder
        New-Item -ItemType Directory -Force -Path $FilePath | Out-Null
        
        #Save XML
        $ExportTemplate.Save("$FilePath\$($FileName).xml")
    }

}

#endregion

#region-[Execution]------------------------------------------------------------
 
    Start-Transcript -Path $sTranscriptFile -Force -Append

    if ((Test-Path $BackupRoot) -eq $false){New-Item -ItemType Directory -Force -Path $BackupRoot | Out-Null}

    Write-Output "Getting list of all scripts."
    
    $ScriptIDs = @{}
    #Query list of all ScriptID's
    if ($($Config.Settings.LastExport) -eq 0) {
        $ScriptIDs = Get-LTData "SELECT ScriptID FROM lt_scripts"
    }
    else{
        $Query = $("SELECT ScriptID FROM lt_scripts WHERE Last_Date > " + "'" + $($Config.Settings.LastExport) +"'")
        $ScriptIDs = Get-LTData $Query   
    }
    
    Write-Output "$(@($ScriptIDs).count) scripts to process."

    #Process each ScriptID
    $n = 0
    foreach ($ScriptID in $ScriptIDs) {
        #Progress bar
        $n++
        Write-Progress -Activity "Backing up LT scripts to $BackupRoot" -Status "Processing ScriptID $($ScriptID.ScriptID)" -PercentComplete  ($n / @($ScriptIDs).count*100)
        
        #Export current script
        Export-LTScript -ScriptID $($ScriptID.ScriptID)
    }

    Write-Output "Export finished."
    $Config.Settings.LastExport = "$($Date.ToString("yyy-MM-dd HH:mm:ss"))"
    $Config.Save("$PSScriptRoot\LT-ScriptExport-Config.xml")

    Stop-Transcript

    #Limit Log file to 50000 lines
    (Get-Content $sTranscriptFile -tail 50000 -readcount 0) | Set-Content $sTranscriptFile -Force -Encoding Ascii

#endregion
