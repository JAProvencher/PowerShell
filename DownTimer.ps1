Param (
    $Start, # time started as HHmm
    $Worked, # HH:MM from WorkBrain for today
    $Total # Total decimal hours worked from WorkBrain
)

$Expected = 8 * (Get-Date).DayOfWeek
$Start = [DateTime]::ParseExact($Start,"HHmm",$null)
$WorkedDec = [math]::Round([timespan]::Parse($Worked).TotalHours,2)
$Total = $Total - $WorkedDec


While ($true) {
    $wrked = [math]::Round(([datetime]::Now - $Start).TotalHours,2)
    $DayWrked = $WorkedDec + $wrked
    $DayPct = $Daywrked / 8

    $W2DPct = ($DayWrked + $Total) / $Expected
    $WeekPct = ($DayWrked + $Total) / 40

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
