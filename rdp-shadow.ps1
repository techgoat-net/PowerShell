# Parameter-Definitionen
param(
    [Parameter(Mandatory=$false,Position=0)]
    [string]$Username="",
    [Parameter(Mandatory=$false,Position=1)]
    [string]$ComputerName=$env:Computername
)
# Fuer Mitteilungen als Dialogfenster fuegen wir den Typ hinzu
Add-Type -AssemblyName System.Windows.Forms
# Aus qwinsta-Ausgabe Objekte machen und diese in ein Array zusammenfassen
$Sessions=@()
(qwinsta /server:$ComputerName) | foreach {
    $objSession=New-Object System.Object
    $objSession | Add-Member -MemberType NoteProperty -Name Server -Value $ComputerName
    $objSession | Add-Member -MemberType NoteProperty -Name SessionID -Value $_.substring(40,8).trim()
    $objSession | Add-Member -MemberType NoteProperty -Name Type -Value $_.substring(1,3).trim()
    $objSession | Add-Member -MemberType NoteProperty -Name Username -Value $_.substring(19,22).trim()
    $objSession | Add-Member -MemberType NoteProperty -Name Status -Value $_.substring(48,8).trim()
    if($objSession.Username -ne "" -and $objSession.Username -ne "USERNAME") {
        $Sessions+=$objSession
    }
}
# Falls kein Benutzername per Parameter übergeben wurde, wird er hier erfragt (und die Liste aller vorhandenen ausgegeben)
if($Username -eq "") {
    $Sessions | Sort-Object { $_.Username } | ft Username,Status,Type -autosize
    $Username=Read-Host -Prompt "Please enter username to shadow on $ComputerName"
}
$found=$false
# Finde die Sitzung des eingegebenen Benutzernamens
$Sessions | Where-Object { $_.Username -eq $Username } | Foreach-Object {
    $found=$true
    # Und spiegel die Sitzung, sofern sie aktiv ist.
    if($_.Status -eq "Active") {
        $sid=$_.SessionID
        iex "& mstsc /shadow:$sid /noconsentprompt"
    }
    # Ansonsten entspr. Ausgabe tätigen.
    else {
        [System.Windows.Forms.MessageBox]::Show("Session of user '$Username' on $ComputerName is inactive.","Information",0,[System.Windows.Forms.MessageBoxIcon]::Information) >$null
    }
}
# Und wenn der Benutzername in der Liste nicht vorkommt, auhc entspr. Meldung ausgeben
if($found -eq $false) {
    [System.Windows.Forms.MessageBox]::Show("User '$Username' seems to be not logged on to $ComputerName.","Information",0,[System.Windows.Forms.MessageBoxIcon]::Information) >$null
}
