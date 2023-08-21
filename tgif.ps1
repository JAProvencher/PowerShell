#get Friday 5:00PM for the current week
$now=Get-Date
[datetime]$TGIF="{0:MM/dd/yyyy} 5:00:00 PM" -f (($now.AddDays( 5 - [int]$now.DayOfWeek)) )
if ((get-date) -ge $TGIF) {
 write-host "TGIF has started without you!" -fore Green
 }
else {
   write-host "Countdown: $( ($tgif -(get-date)).ToString()  )" -fore magenta
 }
 #end script