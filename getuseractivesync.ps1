function colorred {process { Write-Host $_ -ForegroundColor Red }}
function colorgreen {process { Write-Host $_ -ForegroundColor Green }}
function coloryellow {process { Write-Host $_ -ForegroundColor Yellow }}

function sessionTable {
	$activeSessions = Get-ActiveSyncDevice -Mailbox $mailbox | measure
	$sessions = $activeSessions.Count
	echo ""
	Write-Output "Active sessions: $sessions" | colorgreen
	if ($sessions -ne "0") {
		Get-ActiveSyncDevice -Mailbox $mailbox | select @{n="User/Device";e={$_.name}}, @{n="Guid";e={$_.guid}} , @{n="First sync";e={$_.firstsynctime}} | format-table -AutoSize
	} else {
		break
	}
}

$mailbox		= Read-Host 'User/Email'

if (Get-Mailbox $mailbox -ErrorAction SilentlyContinue) {
	sessionTable
} else {
	Write-Output "$mailbox not found." | colorred
	break
}
$deleteSession	= Read-Host 'Delete session? (y/n)?'

if ($deleteSession -eq "y") {
	$session 	= Read-Host 'Delete session (enter guid)'
	if (Get-ActiveSyncDevice -Identity $session -ErrorAction SilentlyContinue) {
		Remove-ActiveSyncDevice -Identity $session
		echo ""
		sessionTable
	} else {
		Write-Output "Session $session not found." | colorred
		break
	}
} else {
	break
}