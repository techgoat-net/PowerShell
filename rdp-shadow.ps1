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
    if($objSession.Username -ne "" -and $objSession.Username -ne "USERNAME" -and $objSesson.Username -notlike $env:username -and $objSession.Status -eq "Active") {
        $Sessions+=$objSession
    }
}
# Wir definieren eine globale Variable fuer den Benutzernamen wie folgt:
$global:Username=$Username # Das brauchen wir naemlich fuer ein eigenes Dialogfenster.
# Falls kein Benutzername per Parameter übergeben wurde, wird er hier erfragt (Buttons)
if($Username -eq "") {
    # Wir laden die Assemblies fuer eigene Formen / Fenster
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    
    # Das Dialogfenster wird definiert
    $objForm=New-Object System.Windows.Forms.Form
    $objForm.Text="Select session to shadow"
    $w=220
    if($Sessions.Count -ge 3) {
        $w=630
    }
    elseif($Sessions.Count -eq 2) {
        $w=430
    }
    $x=0
    $y=0
    # Und wir gehen nun die Sitzungen durch und erstellen jeweils einen Button an entspr. Position
    $Sessions | Sort-Object { $_.Username } | Foreach-Object {
        if($x -ge 3) {
            $y++
            $x=0
        }
        $objBtn=New-Object System.Windows.Forms.Button
        $objBtn.Location=New-Object System.Drawing.Size($(10+$x*200),$(10+$y*34))
        $objBtn.Size=New-Object System.Drawing.Size(190,24)
        $objBtn.Text=$_.Username
        # Und wenn man drauf klickt, wird die globale Variable beschrieben
        $objBtn.Add_Click({$global:Username=$this.Text;$objForm.Close()})
        $objForm.Controls.Add($objBtn)
        $x++
    }
    # Positionierung und Groesse des Dialogfensters
    $objForm.Size=New-Object System.Drawing.Size($w,$(80+$y*34))
    $objForm.StartPosition="CenterScreen"
    # Dialogfenster anzeigen
    [void] $objForm.ShowDialog()
}
$found=$false
# Finde die Sitzung des ausgewaehlten Benutzernamens
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
if($found -eq $false -and $global:Username -ne "") {
    [System.Windows.Forms.MessageBox]::Show("User '$Username' seems to be not logged on to $ComputerName.","Information",0,[System.Windows.Forms.MessageBoxIcon]::Information) >$null
}
