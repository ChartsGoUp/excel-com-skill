param(
    [Parameter(Mandatory=$true)][string]$Python,
    [Parameter(Mandatory=$true)][string]$Runner,
    [Parameter(Mandatory=$true)][string]$Request
)
$ErrorActionPreference='Stop'
$task='OfficeArtifactCOM-'+[guid]::NewGuid().ToString('N')
$registered=$false
try {
    $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
    $user=(Get-CimInstance Win32_ComputerSystem).UserName
    if (!$user) { throw 'No interactive desktop user is signed in.' }
    $account=New-Object Security.Principal.NTAccount($user)
    $sid=$account.Translate([Security.Principal.SecurityIdentifier]).Value
    if ($sid -ne $identity.User.Value) { throw 'Desktop user differs from caller. Refusing cross-user execution.' }
    foreach($path in @($Python,$Runner,$Request)) {
        if($path.Contains('"') -or !(Test-Path -LiteralPath $path)){throw 'Invalid runner path.'}
    }
    # Python's private temp directory can be owned by Administrators when WSL
    # starts an elevated Windows process. The limited desktop token then cannot
    # access the request or even use the directory as its working directory.
    # Grant only this same user's SID access, keeping the directory private.
    $jobDirectory=Split-Path -Parent $Request
    $acl=Get-Acl -LiteralPath $jobDirectory
    $rule=New-Object Security.AccessControl.FileSystemAccessRule($identity.User,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
    $acl.SetAccessRule($rule)
    Set-Acl -LiteralPath $jobDirectory -AclObject $acl
    $arguments='"'+$Runner+'" --job "'+$Request+'"'
    $action=New-ScheduledTaskAction -Execute $Python -Argument $arguments -WorkingDirectory (Split-Path -Parent $Request)
    $principal=New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
    $settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName $task -Action $action -Principal $principal -Settings $settings | Out-Null
    $registered=$true
    Start-ScheduledTask -TaskName $task
    $result=Join-Path (Split-Path -Parent $Request) 'result.json'
    $deadline=(Get-Date).AddSeconds(290)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while(!(Test-Path -LiteralPath $result)) {
        if((Get-Date) -ge $deadline){throw 'Desktop Office worker timed out.'}
        Start-Sleep -Milliseconds 500
        if($sw.Elapsed.TotalSeconds -gt 5) {
            if(!(Test-Path -LiteralPath $result)) {
                $info=Get-ScheduledTaskInfo -TaskName $task
                if($info.LastRunTime -gt (Get-Date).AddMinutes(-5) -and (Get-ScheduledTask -TaskName $task).State -eq 'Ready') {
                    throw ('Desktop worker exited without a result. Exit code: '+$info.LastTaskResult)
                }
            }
        }
    }
    # Result is written atomically; allow the Python wrapper to exit normally.
    for($i=0;$i -lt 10 -and (Get-ScheduledTask -TaskName $task).State -eq 'Running';$i++){Start-Sleep -Milliseconds 200}
} finally {
    if($registered) {
        if((Get-ScheduledTask -TaskName $task).State -eq 'Running'){Stop-ScheduledTask -TaskName $task}
        Unregister-ScheduledTask -TaskName $task -Confirm:$false
    }
}
