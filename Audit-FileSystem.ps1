Function Get-PathData {
    Param([parameter(valuefrompipeline = $true)]$Item)
    Begin {}
    Process {
        if ($Item.endswith("*")) {
            Try { Get-Childitem -Path $Item -Recurse -Force -ErrorAction Stop } Catch { $NotFound.Add("$Item does not exist") }
        }
        else {
            Try { Get-Item -Path $Item -ErrorAction Stop } Catch { $NotFound.Add("$Item does not exist") }
        }
    }
}

$Params = @{
    Property = @{Name = "Path"
        Expression    = { $_.Fullname }
    },
    @{Name         = "Last Modified"
        Expression = { $_.LastWriteTime }
    },
    @{Name         = "Type"
        Expression = {
            switch ($_.getType().name) {
                'FileInfo' { "File" }
                'DirectoryInfo' { "Folder" }
            }
        }
    }
}
$NotFound = [System.Collections.ArrayList]::New()

Foreach ($Server in $Servers) {
    if (-not (Test-NetConnection -Computername $Server -CommonTCPPort SMB -InformationLevel Quiet)) {
        $NotFound.Add("Server $Server cound not be found or is inaccessible.")
        Continue
    }
    Foreach ($Path in $Paths) {
        $ToCheck = "\\$Server\$Path"
        $ToCheck | Get-PathData | Sort-Object -Property Path | Select-Object @Params
    }

}