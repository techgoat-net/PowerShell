###########################################
# Windows Server 2012 R2+ tsadmin PS tool #
# Version 0.7 - written by Nico Domagalla #
# Date: 2018-06-12                        #
###########################################
$ErrorActionPreference="SilentlyContinue"

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
$global:SrvFile=Import-Csv "$global:FilePath\Servers.csv" -Delimiter ";"
$global:TempDir="C:\Temp"
$global:Servers=@()
$global:Sessions=@()
$global:TempSes=@()
$global:TempSrv=@()
$global:tf=""
do {
    $r=Get-Random -Maximum 256 -Minimum 0
    $global:tf="$global:TempDir\$r"
}
while(Test-Path $global:tf)
mkdir $global:tf >$null 2>$null

# This new script block will help to retrieve data asynchronously.
$global:ProcessTimeoutFunc={
    param($srv,$op)
    $p=Start-Process -FilePath "qwinsta" -ArgumentList "/server:$srv" -RedirectStandardOutput $op -Passthru -NoNewWindow
    $to=$null
    $p | Wait-Process -Timeout 10000 -ea 0 -ev $to
}
# A workaround in order to catch users with umlauts and other spec chars as well (Yes, it is an odd function...)
function ConvUmlauts($str) {
    return $str.Replace("`”","ö").Replace("á","ß").Replace(" ","á")
}
# This function will look up sessions of a given server...
function AddSessions($f,$srv) {
    $r=Get-Content $f
    if($r -ne $false) {
        $r | Where {$_ -ne ""} | foreach {
            $usr=ConvUmlauts($_.substring(19,22).trim().ToLower())
            if($usr -ne "" -and $usr -ne "USERNAME" -and $usr -ne $env:username) {
                $ad=Get-ADUser $usr -ErrorAction SilentlyContinue -ErrorVariable $null | Select SamAccountName,SurName,GivenName
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

# This function will display information in status bar
function RefreshStatusBar($SrvListBox,$SesListBox,$StatusBar) {
    if($SesListBox -is [System.Windows.Forms.DataGridView] -and $StatusBar -is [System.Windows.Forms.StatusBar]) {
        $srvcnt=0
        $srvsel=0
        $srvtxt=""
        $sestxt=""
        $delim=""
        if($SesListBox.Rows.Count -gt 0) {
            $sescnt=$SesListBox.Rows.Count
            $sessel=$SesListBox.SelectedRows.Count
            $sesact=($SesListBox.Rows | Where-Object { $_.Cells[4].Value -eq "Active" }).Count
            $sesina=($SesListBox.Rows | Where-Object { $_.Cells[4].Value -ne "Active" }).Count
            $sestxt="$sescnt Sessions ($sessel selected, $sesact active, $sesina inactive)"
        }
        if($SrvListBox -is [System.Windows.Forms.ListBox]) {
            $srvcnt=$SrvListBox.Items.Count
            $srvsel=$SrvListBox.SelectedItems.Count
            $srvtxt="$srvcnt Servers ($srvsel selected)"
        }
        if($srvtxt -ne "" -and $sestxt -ne "") {
            $delim=" - "
        }
        $StatusBar.Text="$srvtxt$delim$sestxt"
    }
}

# These functions will fill the session list view
function RefreshSessions($SrvListBox,$SesListBox,$StatusBar=$null,$pattern=$null) {
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
                    elseif(($global:Servers | Where-Object {$_.Name -eq $srv}).AllowShadow -ne $true) {
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
        if($StatusBar -is [System.Windows.Forms.StatusBar]) {
            RefreshStatusBar $SrvListBox $SesListBox $StatusBar
        }
    }
}

# This function will search for a pattern in listed sessions
function SearchSessions($SesListBox,$StatusBar,$pattern) {
    if($SesListBox -is [System.Windows.Forms.DataGridView]) {
        $sesListBox.Rows.Clear()
        $global:Sessions | Where-Object {$_.Username -like "*$pattern*" -or $_.FirstName -like "*$pattern*" -or $_.LastName -like "*$pattern*"} | Foreach-Object {
            $srv=$_.Server
            $usr=$_.Username
            $r=@($usr,$_.FirstName,$_.LastName,$srv,$_.Status)
            $SesListBox.Rows.Add($r)
            if($_.Status -ne "Active") {
                Foreach($c in $SesListBox.Rows[$($SesListBox.Rows.Count-1)].Cells) {
                    $c.Style.BackColor=[System.Drawing.Color]::FromArgb(255,255,224,224)
                }
            }
            elseif(($global:Servers | Where-Object {$_.Name -eq $srv}).AllowShadow -ne $true) {
                Foreach($c in $SesListBox.Rows[$($SesListBox.Rows.Count-1)].Cells) {
                    $c.Style.BackColor=[System.Drawing.Color]::FromArgb(255,224,224,255)
                }
            }
        }
        if($StatusBar -is [System.Windows.Forms.StatusBar]) {
            RefreshStatusBar $null $SesListBox $StatusBar
        }
    }
}

# This function will refresh the server list view and the session list view.
function RefreshServers($TabControl,$SesListBox,$StatusBar=$null) {
    if($TabControl -is [System.Windows.Forms.TabControl]) {
        for($i=0;$i -lt $($TabControl.TabCount-1);$i++) {
            $SrvListBox=$TabControl.TabPages.Item($i).Controls[0]
            $SrvListBox.Remove_SelectedIndexChanged($global:SelectionChangedFunc)
            $idx=@()
            for($z=0;$z -lt $SrvListBox.Items.Count;$z++) {
                if($SrvListBox.GetSelected($z)) {
                    $idx+=$z
                }
            }
            $SrvListBox.Items.Clear()
            Foreach($srv in ($global:Servers | Where-Object {$_.TabIndex -eq $i} | Sort-Object {$_.Name}).Name) {
                $cnta=($global:Sessions | where-object {$_.Server -eq $srv -and $_.Status -eq "Active"}).Count
                $cntd=($global:Sessions | where-object {$_.Server -eq $srv -and $_.Status -ne "Active"}).Count
                [void]$SrvListBox.Items.Add("$srv [ $cnta | $cntd ]")
            }
            Foreach ($z in $idx) {
                $SrvListBox.SetSelected($z,$true)
            }
            if($i -eq $TabControl.SelectedIndex -and $SesListBox -is [System.Windows.Forms.DataGridView]) {
                RefreshSessions $SrvListBox $SesListBox $StatusBar
            }
            $SrvListBox.Add_SelectedIndexChanged($global:SelectionChangedFunc)
        }
    }
}

# This function will read all sessions...
function RefreshData($Timer) {
	if(Test-Path $global:tf) {
	}
	else {
		mkdir $global:tf >$null 2>$null
	}
	if(Test-Path $global:tf) {
		if($Timer -is [System.Windows.Forms.Timer]) {
			$e=$Timer.Enabled
			$Timer.Enabled=$false
		}
		$global:TempSrv=@()
		$global:TempSes=@()
		$tabs=$global:SrvFile | Group-Object { $_.Grp } | Sort-Object { $_.Name }
		Write-Host "Reading servers from list..."
		for($i=0;$i -lt $tabs.Count;$i++) {
			$global:SrvFile | Where-Object { $_.Grp -eq $tabs[$i].Name } | Foreach-Object {
				$srv=$_.Name
				$f="$global:tf\$srv.txt"
				set-content $f ""
				$shadow=$false
				if($_.Shadow -eq 1) {
					$shadow=$true
				}
				$objServer=New-Object System.Object
				$objServer | Add-Member -MemberType NoteProperty -Name Name -Value $srv
				$objServer | Add-Member -MemberType NoteProperty -Name TabIndex -Value $i
				$objServer | Add-Member -MemberType NoteProperty -Name AllowShadow -Value $shadow
				$global:TempSrv+=$objServer
				Start-Job -ScriptBlock $global:ProcessTimeoutFunc -ArgumentList @($srv,$f) -Name "qwinsta_$srv" >$null 2>$null
			}
		}
		Write-Host "Waiting for $((Get-Job | Where-Object { $_.Name -like `"qwinsta_*`" -and $_.State -eq `"Running`"}).Count) background jobs to finish..."
		while((Get-Job | Where-Object { $_.Name -like "qwinsta_*" -and $_.State -eq "Running"}).Count -gt 0) {
			write-host "#" -nonewline
			Start-Sleep -Seconds 5
		}
		Write-host "`r`nRetrieving sessions..."
		$global:TempSrv | Foreach-Object {
			$srv=$_.Name
			$f="$global:tf\$srv.txt"
			if(Test-Path $f) {
				AddSessions $f $srv
				del $f 2>$null
				if(Test-Path $f) {
					write-host "$srv timed out."
					Remove-Job -Name "qwinsta_$srv" -Force -Confirm:$false
					del $f 2>$null
				}
			}
		}
		$global:Servers=$global:TempSrv
		$global:Sessions=$global:TempSes
		write-host "Finished."
		if($Timer -is [System.Windows.Forms.Timer]) {
			$Timer.Enabled=$e
		}
	}
	else {
		write-host "Failed to create '$global:tf'. Check permissions!"
	}
}

# First load of data
RefreshData

# Now to the form and all that stuff
$objForm=new-Object System.Windows.Forms.Form
$objForm.Icon=[System.IconExtractor]::Extract("shell32.dll",89,$false)
$objForm.Text="TS Admin 2012 by Nico Domagalla"
$objForm.SizeGripStyle="Show"
$objForm.StartPosition="CenterScreen"
$objForm.MinimumSize=New-Object System.Drawing.Size(800,320)
$objForm.Size=New-Object System.Drawing.Size(800,520)

$objStatBar=New-Object System.Windows.Forms.StatusBar
$objForm.Controls.Add($objStatBar)

# A tab view for different server types, and search tab
$objSrvTab=New-Object System.Windows.Forms.TabControl
$objSrvTab.Appearance="Normal"
$objSrvTab.Alignment="Left"
$global:SrvFile | Group-Object { $_.Grp } | Sort-Object { $_.Name } | Foreach-Object {
    $objSrvTab.TabPages.Add($_.Name);
}
$objSrvTab.TabPages.Add("Search")
$objSrvTab.Add_SelectedIndexChanged({
    $objStatBar.Text=""
    if($this.SelectedIndex -lt $($this.TabCount-1)) {
        RefreshSessions $this.TabPages.Item($this.SelectedIndex).Controls[0] $objSesLst $objStatBar
    } else {
        if($objSearchTxt.Text.Trim() -eq "") {
            $objSesLst.Rows.Clear()
        }
        else {
            SearchSessions $objSesLst $objStatBar $objSearchTxt.Text.Trim()
        }
        $objSearchTxt.Focus()
    }
})
$objForm.Controls.Add($objSrvTab)

# These are the lists of servers
$global:SelectionChangedFunc={
    RefreshSessions $this $objSesLst $objStatBar
    $objSesLst.ClearSelection()
    $objDisBtn.Enabled=$false
    $objMirBtn.Enabled=$false
    $objMsgBox.Enabled=$false
    $objMsgBtn.Enabled=$false
}
for($i=0;$i -lt $objSrvTab.TabCount-1;$i++) {
    $objSrvLst=New-Object System.Windows.Forms.ListBox
    $objSrvLst.Location=New-Object System.Drawing.Size(4,4)
    $objSrvLst.Size=New-Object System.Drawing.Size($($objSrvTab.ClientSize.Width-8),$($objSrvTab.ClientSize.Height-8))
    $objSrvLst.BorderStyle="FixedSingle"
    $objSrvLst.SelectionMode="MultiExtended"
    $objSrvLst.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $objSrvLst.Add_SelectedIndexChanged($global:SelectionChangedFunc)
    $objSrvTab.TabPages.Item($i).Controls.Add($objSrvLst)
}

# And the search thing...
$objSearchLbl=New-Object System.Windows.Forms.Label
$objSearchLbl.Location=New-Object System.Drawing.Size(4,4)
$objSearchLbl.AutoSize=$true
$objSearchLbl.Text="Search pattern:"
$objSrvTab.TabPages.Item($objSrvTab.TabCount-1).Controls.Add($objSearchLbl)

$objSearchTxt=New-Object System.Windows.Forms.TextBox
$objSearchTxt.Location=New-Object System.Drawing.Size(4,21)
$objSearchTxt.Width=$($objSrvTab.ClientSize.Width-8)
$objSearchTxt.BorderStyle="FixedSingle"
$objSearchTxt.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$objSearchTxt.Add_KeyDown({
    if($_.KeyCode -eq "Enter" -and $this.Text.Trim() -ne "") {
        SearchSessions $objSesLst $objStatBar $this.Text.Trim()
    }
})
$objSrvTab.TabPages.Item($objSrvTab.TabCount-1).Controls.Add($objSearchTxt)

$objSearchBtn=New-Object System.Windows.Forms.Button
$objSearchBtn.Location=New-Object System.Drawing.Size(4,$($objSearchTxt.Top+$objSearchTxt.Height+4))
$objSearchBtn.Size=New-Object System.Drawing.Size(120,24)
$objSearchBtn.Text="Search"
$objSearchBtn.Enabled=$true
$objSearchBtn.FlatStyle="Flat"
$objSearchBtn.Add_Click({
    SearchSessions $objSesLst $objStatBar $objSearchTxt.Text.Trim()
})
$objSrvTab.TabPages.Item($objSrvTab.TabCount-1).Controls.Add($objSearchBtn)

# This is the session listing grid
$objSesLst=New-Object System.Windows.Forms.DataGridView
$objSesLst.ColumnCount=5
$objSesLst.SelectionMode="FullRowSelect"
$objSesLst.BorderStyle="FixedSingle"
$objSesLst.ReadOnly=$true
$objSesLst.AllowUserToAddRows=$false
$objSesLst.AllowUserToDeleteRows=$false
$objSesLst.AllowUserToResizeRows=$false
$objSesLst.AllowUserToOrderColumns=$true
$objSesLst.AllowUserToResizeColumns=$true
$objSesLst.ColumnHeadersHeightSizeMode="AutoSize"
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
        if(($global:Servers | Where-Object {$_.Name -eq $row.Cells[3].Value}).AllowShadow -ne $true) {
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
    if($objSrvTab.TabPages.Item($objSrvTab.SelectedIndex).Controls[0] -is [System.Windows.Forms.ListBox]) {
        RefreshStatusBar $objSrvTab.TabPages.Item($objSrvTab.SelectedIndex).Controls[0] $this $objStatBar
    }
    else {
        RefreshStatusBar $null $this $objStatBar
    }
})
$objForm.Controls.Add($objSesLst)

# And now the control elements: Buttons, message box...
$RefreshFunc={
    $this.Enabled=$false
    RefreshData $objTimer
    RefreshServers $objSrvTab $objSesLst $objStatBar
    $this.Enabled=$true
}
$objRefBtn=New-Object System.Windows.Forms.Button
$objRefBtn.Text="Refresh"
$objRefBtn.Enabled=$true
$objRefBtn.FlatStyle="Flat"
$objRefBtn.Add_Click($RefreshFunc)
$objForm.Controls.Add($objRefBtn)

$objDisBtn=New-Object System.Windows.Forms.Button
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
        RefreshServers $objSrvTab $objSesLst $objStatBar
    }
})
$objForm.Controls.Add($objDisBtn)

$objMirBtn=New-Object System.Windows.Forms.Button
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

# A timer to automatically refresh (currently disabled)
$objTimer=New-Object System.Windows.Forms.Timer
$objTimer.Interval=600000
#$objTimer.Enabled=$true
$objTimer.Enabled=$false
$objTimer.Add_Tick($RefreshFunc)

# A nice looking function (visuability is everything...)
$objForm.Add_Shown({
    $objSrvTab.Location=New-Object System.Drawing.Size(4,4)
    $objSrvTab.Size=New-Object System.Drawing.Size(200,450)
    $objSrvTab.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $objSesLst.Location=New-Object System.Drawing.Size($($objSrvTab.Left+$objSrvTab.Width+4),$objSrvTab.Top)
    $objSesLst.Size=New-Object System.Drawing.Size(450,$objSrvTab.Height)
    $objSesLst.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $objRefBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$objSesLst.Top)
    $objRefBtn.Size=New-Object System.Drawing.Size(118,24)
    $objRefBtn.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $objDisBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objRefBtn.Top+$objRefBtn.Height+4))
    $objDisBtn.Size=New-Object System.Drawing.Size($objRefBtn.Width,24)
    $objDisBtn.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $objMirBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objDisBtn.Top+$objDisBtn.Height+4))
    $objMirBtn.Size=New-Object System.Drawing.Size($objRefBtn.Width,24)
    $objMirBtn.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $objMsgBox.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objMirBtn.Top+$objMirBtn.Height+4))
    $objMsgBox.Size=New-Object System.Drawing.Size($objRefBtn.Width,$($objSesLst.Height-$objMirBtn.Top-$objMirBtn.Height-28))
    $objMsgBox.Anchor=[System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    $objMsgBtn.Location=New-Object System.Drawing.Size($($objSesLst.Left+$objSesLst.Width+4),$($objMsgBox.Top+$objMsgBox.Height+4))
    $objMsgBtn.Size=New-Object System.Drawing.Size($objRefBtn.Width,24)
    $objMsgBtn.Anchor=[System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
    RefreshServers $objSrvTab $objSesLst $objStatBar
})

#Show the form
[void]$objForm.ShowDialog()
$objTimer.Dispose()
Remove-Item $global:tf -Recurse -Force 2>$null
