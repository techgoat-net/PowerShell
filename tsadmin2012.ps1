###########################################
# Windows Server 2012 R2+ tsadmin PS tool #
# Version 0.2 - written by Nico Domagalla #
# Date: 2018-01-12                        #
###########################################

# We load some libraries for drawing forms...
Add-Type -AssemblyName System.Windows.Forms
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# And we want to have a nice icon from shell32.dll, so we need this part of code (C) Kazun
# https://social.technet.microsoft.com/Forums/windowsserver/en-US/16444c7a-ad61-44a7-8c6f-b8d619381a27/using-icons-in-powershell-scripts?forum=winserverpowershell
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

# Now some global vars ...
$global:FilePath="C:\Scripts\terminalserver"
$global:FileNewServers="$global:FilePath\2012servers.txt"
$global:FileOldServers="$global:FilePath\2008servers.txt"
$global:TempDir="C:\Temp"
$global:Servers=@()
$global:Sessions=@()
$global:TempSes=@()
$global:TempSrv=@()

# This function will call a process with a timeout. (Unfortunately not really reliable)
function ProcessTimeout($proc,$arg,$timeout=10000) {
    do {
        $r=Get-Random -Maximum 24000 -Minimum 23
        $tf="$global:TempDir\$r.txt"
    }
    while(Test-Path $tf)
    $p=Start-Process -FilePath $proc -ArgumentList $arg -RedirectStandardOutput $tf -Passthru
    $to=$null
    $p | Wait-Process -Timeout $timeout -ea 0 -ev $to
    $ret=Get-Content $tf
    del $tf
    if($to) {
        $p | kill
        return $false
    }
    return $ret
}
# A workaround in order to catch users with umlauts and other spec chars as well (Yes, it is an odd function...)
function ConvUmlauts($str) {
    return $str.Replace("`”","ö").Replace("á","ß").Replace(" ","á")
}
# This function will look up sessions of a given server and probably with a given pattern to search in usernames...
function AddSessions($srv,$pattern=$false) {
    #$r=ProcessTimeout "cmd.exe" "/C `"qwinsta /server:$srv`"" 2000
    $r=ProcessTimeout "qwinsta" "/server:$srv" 2000
    if($r -ne $false) {
        $r | Where {$_ -ne ""} | foreach {
            $usr=ConvUmlauts($_.substring(19,22).trim().ToLower())
            if($usr -ne "" -and $usr -ne "USERNAME" -and $usr -ne $env:username) {
                $ad=Get-ADUser $usr | Select SamAccountName,SurName,GivenName
                if($pattern -eq $false -or $ad.SamAccountName -like "*$pattern*" -or $ad.Surname -like "*$pattern*" -or $ad.GivenName -like "*$pattern*") {
                    $objSession=New-Object System.Object
                    $objSession | Add-Member -MemberType NoteProperty -Name Server -Value $srv
                    $objSession | Add-Member -MemberType NoteProperty -Name SessionID -Value $_.substring(40,8).trim()
                    $objSession | Add-Member -MemberType NoteProperty -Name Type -Value $_.substring(1,3).trim()
                    $objSession | Add-Member -MemberType NoteProperty -Name Username -Value $usr
                    $objSession | Add-Member -MemberType NoteProperty -Name FirstName -Value $ad.GivenName
                    $objSession | Add-Member -MemberType NoteProperty -Name LastName -Value $ad.Surname
                    $objSession | Add-Member -MemberType NoteProperty -Name Status -Value $_.substring(48,8).trim()
                    $global:TempSes+=$objSession
                }
            }
        }
    }
    else {
        write-host "$srv timed out." -ForegroundColor Red
    }
}

# This function will read servers into array
function AddServers($file,$newsrv=$true) {
    if(Test-Path $file) {
        Get-Content $file | Foreach {
            $objServer=New-Object System.Object
            $objServer | Add-Member -MemberType NoteProperty -Name Name -Value $_
            $objServer | Add-Member -MemberType NoteProperty -Name IsNewServer -Value $newsrv
            $global:TempSrv+=$objServer
        }
    }
}

# These functions will fill the session list view
function RefreshSessions($SrvListBox,$SesListBox) {
    if($SesListBox -is [System.Windows.Forms.DataGridView]) {
        $sesListBox.Rows.Clear()
        if($SrvListBox -is [System.Windows.Forms.ListBox] -and $SrvListBox.SelectedItems.Count -gt 0) {
            $sc=$SesListBox.SortedColumn
            $sd=$SesLiStBox.SortOrder.ToString()
            Foreach($item in $SrvListBox.SelectedItems) {
                $srv=$item.Split(" ")[0]
                $global:Sessions | Where-Object {$_.Server -eq $srv} | Sort-Object {$_.Username} | Foreach-Object {
                    $usr=$_.Username
                    $r=@($usr,$_.FirstName,$_.LastName,$srv,$_.Status)
                    $SesListBox.Rows.Add($r)
                    if($_.Status -ne "Active") {
                        Foreach($c in $SesListBox.Rows[$($SesListBox.Rows.Count-1)].Cells) {
                            $c.Style.BackColor=[System.Drawing.Color]::FromArgb(255,255,224,224)
                        }
                    }
                    elseif(($global:Servers | Where-Object {$_.Name -eq $srv}).IsNewServer -ne $true) {
                        Foreach($c in $SesListBox.Rows[$($SesListBox.Rows.Count-1)].Cells) {
                            $c.Style.BackColor=[System.Drawing.Color]::FromArgb(255,224,224,255)
                        }
                    }
                }
            }
            if($sd -ne "None") {
                $SesListBox.Sort($sc,$sd)
            }
        }
    }
}

# This function will refresh the server list view and the session list view.
function RefreshServers($SrvListBox,$SesListBox) {
    if($SrvListBox -is [System.Windows.Forms.ListBox]) {
        $idx=@()
        $SrvListBox.Remove_SelectedIndexChanged($global:SelectionChangedFunc)
        for($i=0;$i -lt $SrvListBox.Items.Count;$i++) {
            if($SrvListBox.GetSelected($i)) {
                $idx+=$i
            }
        }
        $SrvListBox.Items.Clear()
        Foreach($srv in ($global:Servers | Sort-Object {$_.Name}).Name) {
            $cnta=($global:Sessions | where-object {$_.Server -eq $srv -and $_.Status -eq "Active"}).Count
            $cntd=($global:Sessions | where-object {$_.Server -eq $srv -and $_.Status -ne "Active"}).Count
            [void]$SrvListBox.Items.Add("$srv [ $cnta | $cntd ]")
        }
        Foreach ($i in $idx) {
            $SrvListBox.SetSelected($i,$true)
        }
        if($SesListBox -is [System.Windows.Forms.DataGridView]) {
            RefreshSessions $SrvListBox $SesListBox
        }
        $SrvListBox.Add_SelectedIndexChanged($global:SelectionChangedFunc)
    }
}

# This function will read all sessions...
function RefreshData($Timer) {
    if($Timer -is [System.Windows.Forms.Timer]) {
        $e=$Timer.Enabled
        $Timer.Enabled=$false
    }
    $global:TempSrv=@()
    $global:TempSes=@()
    Write-Host "Reading servers from list..."
    # Reading the 2008 R2 servers (Those you can't shadow from 2012 R2)
    AddServers $global:FileOldServers $false
    # Reading the new servers...
    AddServers $global:FileNewServers
    Write-host "Reading current user sessions..."
    $global:TempSrv | Foreach-Object {
        Write-Host $_.Name
        AddSessions $_.Name
    }
    $global:Servers=$global:TempSrv
    $global:Sessions=$global:TempSes
    write-host "Finished."
    if($Timer -is [System.Windows.Forms.Timer]) {
        $Timer.Enabled=$e
    }
}

# First load of data
RefreshData

# Now to the form and all that stuff
$objForm=new-Object System.Windows.Forms.Form
$objForm.Icon=[System.IconExtractor]::Extract("shell32.dll",89,$false)
$objForm.Text="TS Admin 2012 by Nico Domagalla"
$objForm.AutoSize=$true
$objForm.AutoSizeMode="GrowAndShrink"
$objForm.SizeGripStyle="Show"
$objForm.StartPosition="CenterScreen"
$objForm.MinimumSize=New-Object System.Drawing.Size(800,320)

$objLbl=New-Object System.Windows.Forms.Label
$objLbl.Location=New-Object System.Drawing.Size(4,4)
$objLbl.AutoSize=$true
$objLbl.Text="Server"
$objForm.Controls.Add($objLbl)

# This is the list of servers
$global:SelectionChangedFunc={
    RefreshSessions $this $objSesLst
    $objSesLst.ClearSelection()
    $objDisBtn.Enabled=$false
    $objMirBtn.Enabled=$false
    $objMsgBox.Enabled=$false
    $objMsgBtn.Enabled=$false
}
$objSrvLst=New-Object System.Windows.Forms.ListBox
$objSrvLst.Location=New-Object System.Drawing.Size(4,21)
$objSrvLst.Size=New-Object System.Drawing.Size(160,480)
$objSrvLst.BorderStyle="FixedSingle"
$objSrvLst.SelectionMode="MultiExtended"
$objSrvLst.Add_SelectedIndexChanged($global:SelectionChangedFunc)
$objForm.Controls.Add($objSrvLst)

$objLbl=New-Object System.Windows.Forms.Label
$objLbl.Location=New-Object System.Drawing.Size($($objSrvLst.Left+$objSrvLst.Width+4),4)
$objLbl.AutoSize=$true
$objLbl.Text="Session"
$objForm.Controls.Add($objLbl)

# This is the session listing grid
$objSesLst=New-Object System.Windows.Forms.DataGridView
$objSesLst.Location=New-Object System.Drawing.Size($($objSrvLst.Left+$objSrvLst.Width+4),$objSrvLst.Top)
$objSesLst.Size=New-Object System.Drawing.Size(480,$objSrvLst.Height)
$objSesLst.ColumnCount=5
$objSesLst.SelectionMode="FullRowSelect"
$objSesLst.BorderStyle="FixedSingle"
$objSesLst.ReadOnly=$true
$objSesLst.AllowUserToAddRows=$false
$objSesLst.AllowUserToDeleteRows=$false
$objSesLst.AllowUserToResizeRows=$false
$objSesLst.AllowUserToOrderColumns=$true
$objSesLst.AllowUserToResizeColumns=$true
$objSesLst.RowHeadersVisible=$false
$objSesLst.AutoSizeRowsMode="AllCells"
$objSesLst.AutoSizeColumnsMode="AllCells"
$objSesLst.Columns[0].Name="Username"
$objSesLst.Columns[1].Name="First Name"
$objSesLst.Columns[2].Name="Last Name"
$objSesLst.Columns[3].Name="Server"
$objSesLst.Columns[4].Name="Status"
$objSesLst.Add_SelectionChanged({
    $AllMirrorable=$true
    Foreach($row in $objSesLst.SelectedRows) {
        if(($global:Servers | Where-Object {$_.Name -eq $row.Cells[3].Value}).IsNewServer -ne $true) {
            $AllMirrorable=$false
        }
    }
    $AllMessageable=$($this.SelectedRows.Count -gt 0)
    Foreach($si in $this.SelectedRows) {
        $usr=$si.Cells[0].Value
        $srv=$si.Cells[3].Value
        $global:Sessions | Where-Object {$_.Server -eq $srv -and $_.Username -eq $usr} | Foreach-Object {
            if($_.Status -ne "Active") {
                $AllMirrorable=$false
                $AllMessageable=$false
            }
        }
    }
    $objMsgBox.Enabled=$AllMessageable
    $objMsgBtn.Enabled=$($AllMessageable -and $objMsgBox.Text.Trim() -ne "")
    $objMirBtn.Enabled=$AllMirrorable
    $objDisBtn.Enabled=$($this.SelectedRows.Count -gt 0)
})
$objForm.Controls.Add($objSesLst)

# And now the control elements: Buttons, message box...
$RefreshFunc={
    $this.Enabled=$false
    RefreshData $objTimer
    RefreshServers $objSrvLst $objSesLst
    $this.Enabled=$true
}
$objRefBtn=New-Object System.Windows.Forms.Button
$objRefBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$objSesLst.Top)
$objRefBtn.Size=New-Object System.Drawing.Size(160,24)
$objRefBtn.Text="Refresh"
$objRefBtn.Enabled=$true
$objRefBtn.FlatStyle="Flat"
$objRefBtn.Add_Click($RefreshFunc)
$objForm.Controls.Add($objRefBtn)

$objDisBtn=New-Object System.Windows.Forms.Button
$objDisBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objRefBtn.Top+$objRefBtn.Height+4))
$objDisBtn.Size=New-Object System.Drawing.Size($objRefBtn.Width,24)
$objDisBtn.Text="Logoff"
$objDisBtn.Enabled=$false
$objDisBtn.FlatStyle="Flat"
$objDisBtn.Add_Click({
    $cnt=$objSesLst.SelectedRows.Count
    if([System.Windows.Forms.MessageBox]::Show("Do you really want to log off $cnt sessions?","Confirmation",4,[System.Windows.Forms.MessageBoxIcon]::Question) -eq "Yes") {
        Foreach($row in $objSesLst.SelectedRows) {
            $usr=$row.Cells[0].Value
            $srv=$row.Cells[3].Value
            $ses=$global:Sessions | Where-Object {$_.Server -eq $srv -and $_.Username -eq $usr}
            $sid=$ses.SessionID
            write-host "Logging off $usr from $srv (SessionID $sid)..."
            iex "& logoff $sid /server:$srv"
            $global:TempSes=@()
            $global:Sessions | Where-Object {$_ -ne $ses} | ForEach-Object {
                $global:TempSes+=$_
            }
            $global:Sessions=$global:TempSes
        }
        write-host "Finished."
        RefreshServers $objSrvLst $objSesLst
    }
})
$objForm.Controls.Add($objDisBtn)

$objMirBtn=New-Object System.Windows.Forms.Button
$objMirBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objDisBtn.Top+$objDisBtn.Height+4))
$objMirBtn.Size=New-Object System.Drawing.Size($objRefBtn.Width,24)
$objMirBtn.Text="Shadow"
$objMirBtn.Enabled=$false
$objMirBtn.FlatStyle="Flat"
$objMirBtn.Add_Click({
    $cnt=$objSesLst.SelectedRows.Count
    if($cnt -le 1 -or [System.Windows.Forms.MessageBox]::Show("Do you really want to shadow $cnt sessions at a time?","Confirmation",4,[System.Windows.Forms.MessageBoxIcon]::Question) -eq "Yes") {
        Foreach($row in $objSesLst.SelectedRows) {
            $usr=$row.Cells[0].Value
            $srv=$row.Cells[3].Value
            $sid=($global:Sessions | Where-Object {$_.Server -eq $srv -and $_.Username -eq $usr}).SessionID
            write-host "Shadowing $usr on $srv (SessionID $sid)..."
            iex "& mstsc /v:$srv /shadow:$sid /control"
        }
    }
})
$objForm.Controls.Add($objMirBtn)

$objMsgBox=New-Object System.Windows.Forms.TextBox
$objMsgBox.Multiline=$true
$objMsgBox.ScrollBars="Vertical"
$objMsgBox.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objMirBtn.Top+$objMirBtn.Height+4))
$objMsgBox.Size=New-Object System.Drawing.Size($objRefBtn.Width,$($objForm.ClientSize.Height-$objMirBtn.Top-$objMirBtn.Height-32))
$objMsgBox.Enabled=$false
$objMsgBox.BorderStyle="FixedSingle"
$objMsgBox.Add_KeyDown({
    $objMsgBtn.Enabled=$($this.Text.Trim() -ne "")
})
$objMsgBox.Add_KeyUp({
    $objMsgBtn.Enabled=$($this.Text.Trim() -ne "")
})
$objForm.Controls.Add($objMsgBox)

$objMsgBtn=New-Object System.Windows.Forms.Button
$objMsgBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objMsgBox.Top+$objMsgBox.Height+4))
$objMsgBtn.Size=New-Object System.Drawing.Size($objRefBtn.Width,24)
$objMsgBtn.Text="Send"
$objMsgBtn.Enabled=$false
$objMsgBtn.FlatStyle="Flat"
$objMsgBtn.Add_Click({
    Foreach($row in $objSesLst.SelectedRows) {
        $usr=$row.Cells[0].Value
        $srv=$row.Cells[3].Value
        write-host "Sending message to $usr on $srv..."
        $objMsgBox.Text.Split("`n") | & msg $usr "/server:$srv"
    }
    $objMsgBox.Text=""
    $this.Enabled=$false
    write-host "Finished."
})
$objForm.Controls.Add($objMsgBtn)

# A workaround to disable autosize of the form but keep the current window size...
$w=$objForm.Width
$h=$objForm.Height
$objForm.AutoSize=$false
$objForm.Width=$w
$objForm.Height=$h

# OK, now initially fill the lists with values...
RefreshServers $objSrvLst $objSesLst

# A timer to automatically refresh (currently disabled)
$objTimer=New-Object System.Windows.Forms.Timer
$objTimer.Interval=600000
#$objTimer.Enabled=$true
$objTimer.Enabled=$false
$objTimer.Add_Tick($RefreshFunc)

# A nice looking function (visuability is everything...)
$ResizeFunc={
    $objSrvLst.Height=$this.ClientSize.Height-32
    $objSesLst.Height=$objSrvLst.Height
    $objMsgBox.Height=$objSrvLst.Height-112
    $objMsgBtn.Top=$objMsgBox.Top+$objMsgBox.Height+4
    $objRefBtn.Width=$this.ClientSize.Width-8-$objSesLst.Left-$objSesLst.Width
    $objDisBtn.Width=$objRefBtn.Width
    $objMirBtn.Width=$objRefBtn.Width
    $objMsgBox.Width=$objRefBtn.Width
    $objMsgBtn.Width=$objRefBtn.Width
}
$objForm.Add_Resize($ResizeFunc)
$objForm.Add_Shown($ResizeFunc)

#Show the form
[void]$objForm.ShowDialog()
$objTimer.Dispose()
