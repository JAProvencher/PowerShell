$vms = @{}

Get-VM | Where-Object {$_.PowerState -eq "PoweredOff"} | ForEach-Object {$vms.Add($_.Name, $_)}

Get-VIEvent -Start (Get-Date).AddDays(-14) -Entity $vms.Values -MaxSamples ([int]::MaxValue) | Where-Object {$_ -is [VMware.Vim.VmPoweredOffEvent]} |

Sort-Object -Property CreatedTime -Unique | ForEach-Object {
    $vms.Remove($_.VM.Name)
}
