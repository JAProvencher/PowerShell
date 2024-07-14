function RunJobFromQueue {
    if ( $queue.Count -gt 0) {
        $j = Start-Job -ScriptBlock $scriptBlock -ArgumentList $queue.Dequeue()
        Register-ObjectEvent -InputObject $j -EventName StateChanged -Action {
            RunJobFromQueue
            Unregister-Event $eventsubscriber.SourceIdentifier
            Receive-Job $eventsubscriber.SourceIdentifier
            Remove-Job $eventsubscriber.SourceIdentifier
            Write-Host "$($queue.count) jobs remaining, $((get-Job -state Running).count) Running"
        } | Out-Null
    }
}
