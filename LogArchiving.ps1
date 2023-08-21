<#PSScriptInfo

.VERSION 1.0.7

.GUID 54850a51-66d5-4916-bb0c-4cd538c59054

.AUTHOR Christos Polydorou

.COMPANYNAME 

.COPYRIGHT (c) 2020 Christos Polydorou. All rights reserved. This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the  GNU General Public License for more details. You should have received a copy of the GNU General Public License along with this program. If not, see <http://www.gnu.org/licenses/>

.TAGS Log File, Archive

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<#

.DESCRIPTION 
	Archive Log Files 

#>

#region Script Information
<#
    Script: LogArchiving.ps1
    Author: Christos Polydorou (christos.polydorou@hotmail.com)
    Purpose: This script is used in order to archive log files by adding the files to a 7zip archive.
	Version History: 1.0.6 Added option to report the result of the task on the subject on the message
                     1.0.5 Added progress bars when archiving files and removing old archive files
                     1.0.4 Added error handling when removing old archive files
                     1.0.3 Added warning and error count report and events
                     1.0.2 Added server name in the report
					 1.0.1 Added informational logging and CC and BCC support for email message
                     0.0.2 Added support to remove old archive files and send report via email
					 0.0.1 Initial version
    Notes: 
		If you choose to enable event loggging, make sure that the source you're using is available.

		The IncludeOnlyFileExtension setting overwrites the ExcludeFileExtension.

	Logging Levels: 0  - Logging Disabled
					10 - Log only errors
					20 - Log errors and warnings
					30 - Log errors, warnings and informational messages

	Event IDs:	1: [Infomation] Started processing directory
				2: [Error] Could not find the directory to process
				3: [Error] Exception getting files
                4: [Error] Could not send the report via email
                5: [Error] Could not find 7zip
                6: [Information] Found 7Zip installation
                7: [Information] Configuration file has been successfully red
                8: [Information] Email message succefully sent
                9: [Information] Started processing task
               10: [Error] Could not find log directory
               11: [Error] Could not get the files from the log directory
               12: [Information] Finished processing task
               13: [Information] Adding files to archive
               14: [Information] Files added to archive
               15: [Information] Removing old archives
               16: [Information] Old archives removed
               17: [Information] Sending report via email
               18: [Information] Report sent
               19: [Warning] The number of warnings encountered (when not zero)
               20: [Error] The number of errors encountered (when not zero)
               21: [Error] Failed to remove old archive file

	Sample XML file (Remove PowerShell Comments on actual XML file)

		<configuration>
		  <Task														# Miltiple tasks can be configured in the same configuration file
			  LogFileDirectory = "E:\tmp"							# The directory where the log files are saved
			  LogFileNamePattern = "te*"							# Select the files whose name matches the pattern

			  ArchiveFilePath = "e:\tmp\IISLogsArchive-[DATE].7z"	# The path to the archive file include [DATE] in order to replace it with the date
																	# the file is created.
			  ArchiveFileDateFormat ="yyyyMMddhhmmss"				# The format of the date for the above
        
			  FilesOlderThanSeconds = ""							# Archive files older than the number of seconds
			  FilesOlderThanMinutes = ""							# Archive files older than the number of minutes
			  FilesOlderThanHours = ""								# Archive files older than the number of hours
			  FilesOlderThanDays = ""								# Archive files older than the number of days
																	# The above settings do not apply as added values

			  IncludeOnlyFileExtension = ".txt;.log"				# Include only files with these file extensions
			  ExcludeFileExtension = ""								# Exclude files with these file extensions
																	# If include is used, exclude will be ignored
        
			  FileRename = "0"										# Rename the files to ".processed" after adding them to the archive
			  FileRemove = "0"										# Delete the files after adding them to the archive

              CleanUpArchiveFiles = "1"                             # Remove old archive files (0: No / 1: Yes
              ArchiveFilesOlderThanSeconds = ""                     # Remove archives older than seconds
              ArchiveFilesOlderThanMinutes = "1"                    # Remove archives older than minues
              ArchiveFilesOlderThanHours = ""                       # Remove archives older than hours
              ArchiveFilesOlderThanDays = ""                        # Remove archives older than days

              EmailNotification = "1"                               # Send report via email (0: No / 1: Yes)
              EmailFrom = "test@lab.local"                          # The sender address
              EmailTo = "user1@lab.local;user2@lab.local"           # The recipient addresses seperated by ";"
              EmailSubject = "Archive Report"                       # The subject of the message
              EmailServer = "mail.lab.local"                        # The mail server to use
              EmailServerPort = "587"                               # The mail server port
              EmailTLS = "true"                                     # The mail server TLS setting
              EmailUsername = "Arhive@lab.local"                    # The username for the email account
              EmailPassword = ""                                    # The password for the email account
              EmailSubjectResult = "append"                         # Update the email message subject with the result of the task ("append", "prepend", "")
			/>
		  <LoggingLevel>0</LoggingLevel>							# The logging level (levels describded above)
		  <EventSource>Log Archiving</EventSource>					# The event log source to use
		  <SevenZipPath>C:\7Zip\7z.exe</SevenZipPath>				# The path to the 7Z executable
		</configuration>
#>
#endregion

#region Script Parameters
[cmdletBinding()]

Param
(
	[string]$ConfigurationPath
)
#endregion

#region Script Variables
$errorCount = 0
$warningCount = 0
#endregion

#region Test if configuration file exists
if (-not $ConfigurationPath) {
	$ConfigurationPath = $PSScriptRoot + "\configuration.xml"
}

if ( (Test-Path $ConfigurationPath) -ne $true) {
	$message = "The configuration file was not found (" + $ConfigurationPath + ")"
	throw $message
}
#endregion

#region Read the configuration
try {
	[xml]$Configuration = Get-Content -Path $ConfigurationPath
	$Tasks = $Configuration.configuration.Task
	$LoggingLevel = $Configuration.configuration.LoggingLevel.Trim()
	$EventSource = $Configuration.configuration.EventSource.Trim()
	$7Zip = $Configuration.configuration.SevenZipPath.Trim()

	$message = "Successfully red the configuration."
	Write-Verbose $message
	if ($LoggingLevel -ge 30) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -Message $message -EventId 7
	}
}
catch {
	$message = "Error reading configuration file (" + $ConfigurationPath + ")`n" + $_.Exception
	throw $message
}    
#endregion

#region Test 7Z
if (!(Test-Path -Path $7Zip)) {
	$message = "Could not find 7Zip"
	if ($LoggingLevel -ge 10 ) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Error -EventId 5 -Message $message
	}
	throw $message
}
else {
	$message = "7Zip installation found."
	Write-Verbose $message
	if ($LoggingLevel -ge 30) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 6 -Message $message
	}
}
#endregion

#region Functions

#region EmailNotification
function EmailNotification {
	if (![string]::IsNullOrEmpty($t.EmailNotification)) {
		if ($t.EmailNotification -ne "0") {
			Write-Verbose "Creating email message..."

			#region Create the message
			# Create the mail item
			$emailMessage = New-Object System.Net.Mail.MailMessage
			$emailMessage.From = $t.EmailFrom
			$t.EmailTo.Split(";") | % { $emailMessage.To.Add($_) }
			if (![string]::IsNullOrEmpty($t.EmailCC)) {
				$t.EmailCC.Split(";") | % { $emailMessage.CC.Add($_) }
			}
			if (![string]::IsNullOrEmpty($t.EmailBCC)) {
				$t.EmailBCC.Split(";") | % { $emailMessage.BCC.Add($_) }
			}
			$emailMessage.Subject = $t.EmailSubject

			# Update the subject of the message with the result of the task
			if ($t.EmailSubjectResult.length -gt 0) {
				if ($errorCount -gt 0) {
					$EmailMessageSubjectResult = "Error"
				}
				else {
					if ($warningCount -gt 0) {
						$EmailMessageSubjectResult = "Warning"
					}
					else {
						$EmailMessageSubjectResult = "Success"
					}
				}

				if ($t.EmailSubjectResult.ToLower() -eq "prepend") {
					$emailMessage.Subject = $EmailMessageSubjectResult + " - " + $t.EmailSubject
				}
				else {
					$emailMessage.Subject = $t.EmailSubject + " - " + $EmailMessageSubjectResult 
				}
			}

			$emailMessage.IsBodyHtml = $true

			# Synthesize the body of the message


			[string]$body = "<p><strong>Server info</strong></p>"
			$body += "Server name: $($env:COMPUTERNAME)"
			$body += "<p><strong>Task Configuration</strong></p>"
			$body += "Log File Directory: $($t.LogFileDirectory)<br/>"
			$body += "Log File Extensions Included: $($t.IncludeOnlyFileExtension)<br/>"
			$body += "Log File Extensions Excluded: $($t.ExcludeFileExtension)<br/>"
			$body += "Archive File: $($archivePath)<br/>"
			$body += "Log File Remove: $($t.FileRemove)<br/>"
			$body += "Log File Rename: $($t.FileRename)<br/>"
    
			# Add the files that have -been added to the archive
			$body += "<p><strong>Files Added to the archive</strong></p>"
			$files | % FullName | % { $body += "$_ <br/>" }

			# Add archive files that have been removed
			$body += "<p><strong>Removed Archive Files</strong></p>"
			$archiveFilesToRemove | % { $body += "$_ <br/>" }

			# Add number of 7Zip failures
			$body += "<p><strong>7Zip Failures</strong></p>"
			$body += "7Zip failed $7ZipFailures times."

			$emailMessage.Body = $body 

			# Add the 7Zip diagnostic logging
			$7ZipStandardOutputAttachment = [System.Net.Mail.Attachment]::CreateAttachmentFromString($7ZipStandardOutput, "text/csv")
			$7ZipStandardOutputAttachment.ContentDisposition.FileName = "7ZipStandardOutput.txt"
			$emailMessage.Attachments.Add($7ZipStandardOutputAttachment)

			$7ZipStandardErrorAttachment = [System.Net.Mail.Attachment]::CreateAttachmentFromString($7ZipStandardError, "text/csv")
			$7ZipStandardErrorAttachment.ContentDisposition.FileName = "7ZipStandardError.txt"
			$emailMessage.Attachments.Add($7ZipStandardErrorAttachment)
            
			# Create the SMTP client 
			$SMTPClient = New-Object System.Net.Mail.SmtpClient($t.EmailServer, $t.EmailServerPort)
			$SMTPClient.EnableSsl = $t.EmailTLS

			# Check if credentials should be used
			if (![string]::IsNullOrEmpty($t.EmailUsername)) {
				$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($t.EmailUsername , $t.EmailPassword)
			}
			#endregion
            
			#region Send the message
			try {
				Write-Verbose "Sending email message..."
				$SMTPClient.Send($emailMessage)
				$message = "Email message sent."
				Write-Verbose $message
				if ($LoggingLevel -ge 30) {
					Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 8 -Message $message
				}
			}
			catch {
				$message = "Failed to send the report via email.`n" + $_
				Write-Error $message
				$errorCount++
				if ($LoggingLevel -ge 10 ) {
					Write-EventLog -LogName Application -Source $EventSource -EntryType Error -EventId 4 -Message $message
				}
			}
			#endregion
		}
	}
}
#endregion

#region TestFileNamePattern
function TestFileNamePattern {
	if (![string]::IsNullOrEmpty($logFileNamePattern)) {
		if ($f.Name -like $logFileNamePattern) {
			return $true
		}
		else {
			return $false
		}
	}
	else {
		return $true
	}

}
#endregion

#region TestAgeRestrictions
function TestAgeRestrictions {
	# Check the seconds age
	if (![string]::IsNullOrEmpty($t.FilesOlderThanSeconds)) {
		if ($timespan.TotalSeconds -gt [double]$t.FilesOlderThanSeconds) {
			return $true
		}
	}

	# Check the minutes age
	if (![string]::IsNullOrEmpty($t.FilesOlderThanMinutes)) {
		if ($timespan.TotalMinutes -gt [double]$t.FilesOlderThanMinutes) {
			return $true
		}
	}

	# Check the hours age
	if (![string]::IsNullOrEmpty($t.FilesOlderThanHours)) {
		if ($timespan.TotalHours -gt [double]$t.FilesOlderThanHours) {
			return $true
		}
	}

	# Check the days age
	if (![string]::IsNullOrEmpty($t.FilesOlderThanDays)) {
		if ($timespan.TotalDays -gt [double]$t.FilesOlderThanDays) {
			return $true
		}
	}
	
	# If there are no age restriction return True
	if ([string]::IsNullOrEmpty($t.FilesOlderThanDays) -and [string]::IsNullOrEmpty($t.FilesOlderThanHours) -and [string]::IsNullOrEmpty($t.FilesOlderThanMinutes) -and [string]::IsNullOrEmpty($t.FilesOlderThanSeconds)) {
		return $true
	}
	else {
		return $false
	}
}
#endregion

#region TestArchiveAgeRestrictions
function TestArchiveAgeRestrictions {
	# Check the seconds age
	if (![string]::IsNullOrEmpty($t.ArchiveFilesOlderThanSeconds)) {
		if ($timespan.TotalSeconds -gt [double]$t.ArchiveFilesOlderThanSeconds) {
			return $true
		}
	}

	# Check the minutes age
	if (![string]::IsNullOrEmpty($t.ArchiveFilesOlderThanMinutes)) {
		if ($timespan.TotalMinutes -gt [double]$t.ArchiveFilesOlderThanMinutes) {
			return $true
		}
	}

	# Check the hours age
	if (![string]::IsNullOrEmpty($t.ArchiveFilesOlderThanHours)) {
		if ($timespan.TotalHours -gt [double]$t.ArchiveFilesOlderThanHours) {
			return $true
		}
	}

	# Check the days age
	if (![string]::IsNullOrEmpty($t.ArchiveFilesOlderThanDays)) {
		if ($timespan.TotalDays -gt [double]$t.ArchiveFilesOlderThanDays) {
			return $true
		}
	}
	
	# If there are no age restrictions return True
	if ([string]::IsNullOrEmpty($t.ArchiveFilesOlderThanDays) -and [string]::IsNullOrEmpty($t.ArchiveFilesOlderThanHours) -and [string]::IsNullOrEmpty($t.ArchiveFilesOlderThanMinutes) -and [string]::IsNullOrEmpty($t.ArchiveFilesOlderThanSeconds)) {
		return $false
	}
}
#endregion

#region TestFileExtensionRestrictions
function TestFileExtensionRestrictions {
	if ([string]::IsNullOrEmpty($fileExtension)) {
		# Could not define the extension of the file
		return $false
	}
	else {
		if ([string]::IsNullOrEmpty($fileExtensionsToInclude) -and [string]::IsNullOrEmpty($fileExtensionsToExclude)) {
			return $true
		}
		else {
			if ($fileExtensionsToInclude) {
				# Include the file due to it's extension
				if ($fileExtensionsToInclude -contains $fileExtension) {
					return $true
				}
				else {
					return $false
				}
			}

			if ($fileExtensionsToExclude -notcontains $fileExtension) {
				return $true
			}
		}
	}

	return $false
}
#endregion
#endregion

#region Process each task in the configuration
foreach ($t in $Tasks) {
	# Get the current date
	$now = [datetime]::Now

	# Variables for the report
	[string]$7ZipStandardOutput = [string]::Empty
	[string]$7ZipStandardError = [string]::Empty

	# The number of times 7Zip failed
	[int]$7ZipFailures = 0

	# Parse task parameters
	$logFolder = $t.LogFileDirectory
	$logFileNamePattern = $t.LogFileNamePattern
	$archivePath = $t.ArchiveFilePath
	$archivePathFormat = $t.ArchiveFileDateFormat
	$fileExtensionsToInclude = $t.IncludeOnlyFileExtension.ToLower().Split(';')
	$fileExtensionsToExclude = $t.ExcludeFileExtension.ToLower().Split(';')

	# Start processing the task
	$message = "Started processing directory $logFolder"
	Write-Verbose $message
	if ($LoggingLevel -ge 30 ) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 9 -Message $message
	}    

	# Test if logFolder exists
	if ( (Test-Path -Path $logFolder) -ne $true) {
		$message = "Could not access folder $logFolder"
		if ($LoggingLevel -ge 10) {
			Write-EventLog -LogName Application -Source $EventSource -EntryType Error -EventId 10 -Message $message
		}
		Write-Error $message
		$warningCount++
		continue
	}

	# Get the files
	$files = New-Object -TypeName System.Collections.ArrayList
	Write-Verbose "Getting files on directory."
	try {
		$allFiles = Get-ChildItem -Path $logFolder -File -Force -ErrorAction Stop
	}
	catch {
		$message = "Exception getting the files in the directory."
		Write-Error $message
		$errorCount++
		if ($LoggingLevel -ge 10) {
			Write-EventLog -LogName Application -Source $EventSource -EntryType Error -EventId 11 -Message $message
		}
	}

	#region Form the archive file name
	if ($archivePath -like "*[DATE]*") {
		# We have to add the date to the archive file name
		if (![string]::IsNullOrEmpty($archivePathFormat)) {
			$archivePath = $archivePath.Replace("[DATE]", $now.Tostring($archivePathFormat))
		}
	}
	#endregion

	#region Apply the rules
	foreach ($f in $allFiles) {
		# Get the timespan between now and the files last write time
		$timespan = $now - $f.lastwriteTime

		# Get the extension of the file
		try {
			$fileExtension = $f.Name.ToLower().Substring($f.Name.LastIndexOf('.'))
		}
		catch {
			$fileExtension = $null
		}

		#region Test the age and the extension of the file
		if ($fileExtension -ne $null) {
			if ((TestFileExtensionRestrictions -eq $true) -and (TestAgeRestrictions -eq $true) -and (TestFileNamePattern -eq $true)) {
				Write-Verbose "`t File $($f.FullName) will be added to the archive."
				$files.Add($f) |
				Out-Null
			}
		}
		#endregion
	}
	#endregion

	#region Add each file to the archive and then delete it
	$message = "Adding files to archive."
	Write-Verbose $message
	if ($LoggingLevel -ge 30 ) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 13 -Message $message
	}    

	# Reset the progress
	$progressCounter = 0

	foreach ($f in $files) {
		# Update the progress
		$progressCounter++
		Write-Progress -Activity "Archiving files" -status "File $($f.FullName)" -PercentComplete ($progressCounter / $files.count * 100)
		Write-Verbose "Adding file $($f.Fullname) to archive $archivePath"

		# Create the process start information
		$pinfo = New-Object System.Diagnostics.ProcessStartInfo
		$pinfo.FileName = $7Zip
		$pinfo.UseShellExecute = $false
		$pinfo.CreateNoWindow = $true
		$pinfo.Arguments = ("a", ('"' + $archivePath + '"'), ('"' + $f.FullName + '"'))
		$pinfo.RedirectStandardError = $true
		$pinfo.RedirectStandardOutput = $true
		
		# Start the process
		$p = New-Object System.Diagnostics.Process
		$p.StartInfo = $pinfo
		$p.Start() | Out-Null
		$p.WaitForExit()

		# Save the 7Zip output
		$7ZipStandardOutput += "`n" + $p.StandardOutput.ReadToEnd()
		$7ZipStandardError += "`n" + $p.StandardError.ReadToEnd()

		# Check the return value of the process
		if ($p.ExitCode -ne 0) {
			$7ZipFailures++
			Write-Error ("Could not add file $($f.fullname) to archive $archivePath`n" + $p.StandardError.ReadToEnd())
		}
		else {
			# Rename the file
			if (![string]::IsNullOrEmpty($t.FileRename)) {
				if ($t.FileRename -ne "0") {
					Rename-Item -Path $f.FullName -NewName ($f.Name + ".processed") -Force
				}
			}

			# Remove the file
			if (![string]::IsNullOrEmpty($t.FileRemove)) {
				if ($t.FileRemove -ne "0") {
					Remove-Item -Path $f.FullName -Force
				}
			}
		}
		$p.Close()
	}
	$message = "Finished adding files to archive."
	Write-Verbose $message
	if ($LoggingLevel -ge 30 ) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 14 -Message $message
	}    
	#endregion

	#region Cleanup Archive Files
	# Get the current datetime since the archive task may take some time to complete
	$message = "Removing old archives."
	Write-Verbose $message
	if ($LoggingLevel -ge 30 ) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 15 -Message $message
	}    
	$archiveFilesToRemove = @()
	$now = [datetime]::Now
	if (![string]::IsNullOrEmpty($t.CleanUpArchiveFiles)) {
		if ($t.CleanUpArchiveFiles -ne "0") {
			# Get the list of archive files
			Write-Verbose "Removing old archive files..."
			$archiveFiles = Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($archivePath))

			# Reset the progress
			$progressCounter = 0

			foreach ($af in $archiveFiles) {
				# Update the progress
				$progressCounter++

				$timespan = $now - $af.lastwriteTime

				if (TestArchiveAgeRestrictions) {
					Write-Verbose "`t Removing archive file $($af.FullName)"

					# Display the progress
					Write-Progress -Activity "Removing old archive files" -status "File $($af.FullName)" -PercentComplete ($progressCounter / $archiveFiles.count * 100)
					Write-Verbose "Removing archive file $($af.Fullname)"

					try {
						Remove-Item $af.FullName -Force
						$archiveFilesToRemove += $af.FullName
					}
					catch {
						$message = "Failed to remove archive file $($af.fullname)."
						Write-Verbose $message
						if ($LoggingLevel -ge 10 ) {
							Write-EventLog -LogName Application -Source $EventSource -EntryType Error -EventId 21 -Message $message
						}    
					}
				}
			}
		}
	}
	$message = "Finished removing old archives."
	Write-Verbose $message
	if ($LoggingLevel -ge 30 ) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 16 -Message $message
	}    
	#endregion

	#region send report
	if (![string]::IsNullOrEmpty($t.EmailNotification)) {
		if ($t.EmailNotification -ne 0) {
			$message = "Sending report via Email."
			Write-Verbose $message
			if ($LoggingLevel -ge 30 ) {
				Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 17 -Message $message
			}
			
			EmailNotification

			$message = "Report sent."
			Write-Verbose $message
			if ($LoggingLevel -ge 30 ) {
				Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 18 -Message $message
			}    
		}
	}
	#endregion

	# Finished processing the task
	$message = "Finished processing directory $logFolder"
	Write-Verbose $message
	if ($LoggingLevel -ge 30 ) {
		Write-EventLog -LogName Application -Source $EventSource -EntryType Information -EventId 12 -Message $message
	}

	# Write the warnings count on event viewer
	if ($LoggingLevel -ge 20) {
		if ($warningCount -gt 0) {
			$message = "$warningCount warnings encountered during process."
			Write-EventLog -LogName Application -Source $EventSource -EntryType Warning -EventId 19 -Message $message            
		}
	}

	# Write the errors count on event viewer
	if ($LoggingLevel -ge 10) {
		if ($errorCount -gt 0) {
			$message = "$errorCount errors encountered during process."
			Write-EventLog -LogName Application -Source $EventSource -EntryType Error -EventId 20 -Message $message            
		}
	}

}
#endregion
