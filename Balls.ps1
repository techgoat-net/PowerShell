Add-Type -AssemblyName System.Windows.Forms
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$OldCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
$OldUICulture = [System.Threading.Thread]::CurrentThread.CurrentUICulture
[System.Threading.Thread]::CurrentThread.CurrentCulture = "en-US"
[System.Threading.Thread]::CurrentThread.CurrentUICulture = "en-US"
$global:Controls=@()
function SpawnWindow {
    $objForm=new-object System.Windows.Forms.Form
    $objForm.SizeGripStyle="Show"
    $objForm.MinimumSize=new-object system.drawing.size(120,120)
    $objForm.Size=new-object system.drawing.size(320,240)
    $objForm.StartPosition="CenterScreen"
    $objForm.MinimizeBox=$false
    $objForm.MaximizeBox=$false
    ForEach($C in $global:Controls) {
        $objForm.Controls.Add($C)
    }
    $objTimer=new-object System.Windows.Forms.Timer
    $objTimer.Interval=10
    $objTimer.Enabled=$true
    $objTimer.Add_Tick({
        ForEach($C in $objForm.Controls) {
            if($C -is [System.Windows.Forms.RadioButton]) {
                $xdir=[convert]::ToDouble($C.AccessibleDescription)
                $ydir=[convert]::ToDouble($C.AccessibleName)
                $x=[convert]::ToDouble($C.Name)
                $y=[convert]::ToDouble($C.AccessibleDefaultActionDescription)
                $x+=$xdir
                $y+=$ydir
                if(($x+$C.Width) -gt $objForm.ClientSize.Width) {
                    if($xdir -gt 0) {
                        $xdir*=-1
                        }
                    $x=$objForm.ClientSize.Width-$C.Width
                }
                elseif($x -lt 0 -and $xdir -lt 0) {
                    $xdir*=-1
                    $x=0
                }
                if(($y+$C.Height) -gt $objForm.ClientSize.Height) {
                    if($ydir -gt 0) {
                        $ydir*=-1
                        }
                    $y=$objForm.ClientSize.Height-$C.Height
                }
                elseif($y -lt 0 -and $ydir -lt 0) {
                    $ydir*=-1
                    $y=0
                }
                $C.Location=new-object System.Drawing.Size([math]::Round($x),[math]::Round($y))
                $C.AccessibleDescription=$xdir
                $C.AccessibleName=$ydir
                $C.Name=$x
                $C.AccessibleDefaultActionDescription=$y
            }
        }
    })
    [void] $objForm.ShowDialog()
    ForEach($C in $objForm.Controls) {
        $C.Dispose()
    }
    $objForm.Dispose()
    $objTimer.Dispose()
}
function SpawnBall($PosX,$PosY,$DirX,$DirY) {
    $objRadio=new-object System.Windows.Forms.RadioButton
    $objRadio.Text=""
    $objRadio.Checked=$true
    $objRadio.Location=new-object System.Drawing.Size($PosX,$PosY)
    $objRadio.Autosize=$true
    $objRadio.FlatStyle="Flat"
    $objRadio.AccessibleDescription=$DirX
    $objRadio.AccessibleName=$DirY
    $objRadio.Name=$PosX
    $objRadio.AccessibleDefaultActionDescription=$PosY
    $global:Controls+=$objRadio
}
$BallCount=Get-Random -Minimum 20 -Maximum 50
1..$BallCount | forEach {
    SpawnBall $(Get-Random -Minimum 0 -Maximum 310) $(Get-Random -Minimum 0 -Maximum 230) $((Get-Random -Minimum 100 -Maximum 8000)/1000) $((Get-Random -Minimum 100 -Maximum 8000)/1000)
}
SpawnWindow
[System.Threading.Thread]::CurrentThread.CurrentCulture = $OldCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $OldUICulture
