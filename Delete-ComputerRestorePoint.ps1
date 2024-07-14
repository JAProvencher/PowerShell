function Delete-ComputerRestorePoint{
<#
    .SYNOPSIS
        Function to Delete Windows System Restore points
        
    .DESCRIPTION
		Deletes Windows System Restore point(s) passed as an argument or via pipeline
    
    .PARAMETER RestorePoint
        Restore point(s) to be deleted (retrieved and optionally filtered from Get-ComputerRestorePoint
    
    .EXAMPLE  
	    #use -WhatIf to see what would have happened
	    Get-ComputerRestorePoint | Delete-ComputerRestorePoints -WhatIf
		
	.EXAMPLE  
	    #delete all System Restore Points older than 14 days
	    $removeDate = (Get-Date).AddDays(-14)
	    Get-ComputerRestorePoint | 
		    Where { $_.ConvertToDateTime($_.CreationTime) -lt  $removeDate } | 
		    Delete-ComputerRestorePoints 
		 
#>
	[CmdletBinding(SupportsShouldProcess=$True)]param(  
	    [Parameter(
	        Position=0, 
	        Mandatory=$true, 
	        ValueFromPipeline=$true
		)]
	    $RestorePoint
	)
	begin{
		$fullName="SystemRestore.DeleteRestorePoint"
		#check if the type is already loaded
		$isLoaded=([AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object {$_.GetTypes()} | Where-Object {$_.FullName -eq $fullName}) -ne $null
		if (!$isLoaded){
			$SRClient= Add-Type   -memberDefinition  @"
		    	[DllImport ("Srclient.dll")]
		        public static extern int SRRemoveRestorePoint (int index);
"@  -Name DeleteRestorePoint -NameSpace SystemRestore -PassThru
		}
	}
	process{
		foreach ($restorePoint in $RestorePoint){
			if($PSCmdlet.ShouldProcess("$($restorePoint.Description)","Deleting Restorepoint")) {
		 		[SystemRestore.DeleteRestorePoint]::SRRemoveRestorePoint($restorePoint.SequenceNumber)
			}
		}
	}
}