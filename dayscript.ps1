Param (
    $Start, # time started as HHmm
    $Worked, # HH:MM from WorkBrain for today
    $Total # Total decimal hours worked from WorkBrain
)

$Expected = 8 * (Get-Date).DayOfWeek
$Day = @{
    Start = [DateTime]::ParseExact($Start,"HHmm",$null)
    Worked = $($Worked)
    Total = $Total
}
$Day.WorkedDec = [timespan]::Parse($Day.Worked).TotalHours
$ScreenPos = @{
    X=0
    Y=0
}

Clear-Host

$PSStyle.Progress.View = 'Minimal'
$PSStyle.Progress.MaxWidth = 40


$host.UI.RawUI.CursorPosition = $ScreenPos
$DayEnd = ($Day.Start.AddHours(8 - $Day.WorkedDec).ToShortTimeString())
$W2DEnd = ($Day.Start.AddHours($Expected - $Day.Total)).ToShortTimeString()

$host.UI.RawUI.WindowTitle = "[8]: $DayEnd - [$($Expected)]: $W2DEnd"
While ($true) {
    $Worked = $Day.WorkedDec + ([datetime]::Now - $Day.Start).TotalHours
    $DayPct = ($Day.WorkedDec + ([datetime]::Now - $Day.Start).TotalHours) / 8
    $W2DPct = ((([datetime]::Now - $Day.Start).TotalHours  + $Day.Total) / $Expected)
    $WeekPct = ((([datetime]::Now - $Day.Start).TotalHours + $Day.Total) / 40)
    $ToGo1 = [timespan]::FromHours(8 - ($Day.WorkedDec + ([datetime]::Now - $Day.Start).TotalHours)).tostring("hh\:mm\:ss")
    $ToGo2 = [timespan]::FromHours($Expected - ($Day.Total + ([datetime]::Now - $Day.Start).TotalHours)).tostring("hh\:mm\:ss")
    if ($Expected -eq 40) {
        $host.UI.RawUI.WindowTitle = "[8]: $DayEnd - [$($Expected)]: $ToGo2"
    } else {
        $host.UI.RawUI.WindowTitle = "[8]: $DayEnd - [$($Expected)]: $W2DEnd"
    }
     

    
    if ($DayPct -lt 1) {
        if ($Expected -ne 8) {
            Write-Progress -Activity "Day:" -Status "$($DayPct.toString('P')) of 8 hours" -PercentComplete ((1 - $DayPct) * 100) -Id 0
        }
    } else {
        $over = [timespan]::FromHours(8 - ($day.workeddec + ([datetime]::Now - $Day.Start).TotalHours)).tostring("hh\:mm\:ss")
        Write-Progress -Activity "Over: " -Status "$($over)" -id 0
    }
    if ($w2dPct -lt 1) {
        Write-Progress -Activity "Week to Day" -Status "$($W2DPct.toString('P')) of $expected" -PercentComplete ((1 - $W2DPct) * 100) -Id 1
    }else {
        $ot = [timespan]::FromHours((([datetime]::Now - $Day.STart).TotalHours + $Day.Total) - $Expected).tostring("hh\:mm\:ss")
        if ($Expected -lt 40) {
            Write-Progress -Activity "Toward Friday:" -Status $($ot) -id 1
        } else {
            Write-Progress -Activity "Overtime: " -Status $($ot) -id 1
        }
    }
    if ($Expected -lt 40) {
        Write-Progress -Activity "Week" -Status "$($WeekPct.toString('P')) of 40" -PercentComplete ((1 - $WeekPct) * 100) -Id 2
    }
}
