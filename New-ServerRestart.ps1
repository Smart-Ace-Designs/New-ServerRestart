<#
============================================================================================================================
Script: New-ServerRestart
Author: Smart Ace Designs

Notes:
This script schedules a one time restart on a remote server.
============================================================================================================================
#>

#region Settings
$SupportContact = "Smart Ace Designs"
$TaskAction = "Shutdown.exe"
$TaskArgument = "/r /t 1"
$TaskName = "One-time Restart"
$TaskPath = "CorporationName"
$TaskUser = "System"
#endregion

#region Assemblies
Add-Type -AssemblyName System.Windows.Forms
#endregion

#region Appearance
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Controls
$FormMain = New-Object -TypeName System.Windows.Forms.Form
$GroupBoxMain = New-Object -TypeName System.Windows.Forms.GroupBox
$LabelName = New-Object -TypeName System.Windows.Forms.Label
$TextBoxName = New-Object -TypeName System.Windows.Forms.TextBox
$LabelTime = New-Object -TypeName System.Windows.Forms.Label
$TextBoxTime = New-Object -TypeName System.Windows.Forms.TextBox
$LabelDate = New-Object -TypeName System.Windows.Forms.Label
$DateTimePickerDate = New-Object -TypeName System.Windows.Forms.DateTimePicker
$ButtonRun = New-Object -TypeName System.Windows.Forms.Button
$ButtonClose = New-Object -TypeName System.Windows.Forms.Button
$StatusStripMain = New-Object -TypeName System.Windows.Forms.StatusStrip
$ToolStripStatusLabelMain = New-Object -TypeName System.Windows.Forms.ToolStripStatusLabel
$ErrorProviderMain = New-Object -TypeName System.Windows.Forms.ErrorProvider
#endregion

#region Forms
$ShowFormMain =
{
    $FormWidth = 330
    $FormHeight = 260

    $FormMain.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path)
    $FormMain.Text = "Schedule Server Restart"
    $FormMain.Font = New-Object -TypeName System.Drawing.Font("MS Sans Serif",8)
    $FormMain.ClientSize = New-Object -TypeName System.Drawing.Size($FormWidth,$FormHeight)
    $FormMain.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $FormMain.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $FormMain.MaximizeBox = $false
    $FormMain.AcceptButton = $ButtonRun
    $FormMain.CancelButton = $ButtonClose
    $FormMain.Add_Shown($FormMain_Shown)

    $GroupBoxMain.Location = New-Object -TypeName System.Drawing.Point(10,5)
    $GroupBoxMain.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 20),($FormHeight - 80))
    $FormMain.Controls.Add($GroupBoxMain)

    $LabelName.Location = New-Object -TypeName System.Drawing.Point(15,15)
    $LabelName.AutoSize = $true
    $LabelName.Text = "Server Name:"
    $GroupBoxMain.Controls.Add($LabelName)

    $TextBoxName.Location = New-Object -TypeName System.Drawing.Point(15,35)
    $TextBoxName.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 50),20)
    $TextBoxName.TabIndex = 0
    $TextBoxName.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper
    $TextBoxName.MaxLength = 15
    $TextBoxName.Add_TextChanged($TextBoxName_TextChanged)
    $GroupBoxMain.Controls.Add($TextBoxName)

    $LabelTime.Location = New-Object -TypeName System.Drawing.Point(15,70)
    $LabelTime.AutoSize = $true
    $LabelTime.Text = "Restart Time (24-Hour Format 00:00:00):"
    $GroupBoxMain.Controls.Add($LabelTime) 

    $TextBoxTime.Location = New-Object -TypeName System.Drawing.Point(15,90)
    $TextBoxTime.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 50),20)
    $TextBoxTime.TabIndex = 1
    $TextBoxTime.Text = "03:00:00"
    $GroupBoxMain.Controls.Add($TextBoxTime)

    $LabelDate.Location = New-Object -TypeName System.Drawing.Point(15,125)
    $LabelDate.AutoSize = $true
    $LabelDate.Text = "Restart Date:"
    $GroupBoxMain.Controls.Add($LabelDate) 

    $DateTimePickerDate.Location = New-Object -TypeName System.Drawing.Point(15,145)
    $DateTimePickerDate.Size = New-Object -TypeName System.Drawing.Size(($FormWidth - 50),20)
    $DateTimePickerDate.TabIndex = 2
    $DateTimePickerDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $DateTimePickerDate.CustomFormat = "MM/dd/yyyy"
    $DateTimePickerDate.MinDate = [System.DateTime]::Today
    $DateTimePickerDate.Text = ((Get-Date).AddDays(1)).ToString("MM/dd/yyyy")
    $GroupBoxMain.Controls.Add($DateTimePickerDate)

    $ButtonRun.Location = New-Object -TypeName System.Drawing.Point(($FormWidth - 175),($FormHeight - 60))
    $ButtonRun.Size = New-Object -TypeName System.Drawing.Size(75,25)
    $ButtonRun.TabIndex = 100
    $ButtonRun.Enabled = $false
    $ButtonRun.Text = "Run"
    $ButtonRun.Add_Click($ButtonRun_Click)
    $FormMain.Controls.Add($ButtonRun)

    $ButtonClose.Location = New-Object -TypeName System.Drawing.Point(($FormWidth - 85),($FormHeight - 60))
    $ButtonClose.Size = New-Object -TypeName System.Drawing.Size(75,25)
    $ButtonClose.TabIndex = 101
    $ButtonClose.Text = "Close"
    $ButtonClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $FormMain.Controls.Add($ButtonClose)

    $StatusStripMain.SizingGrip = $false
    $StatusStripMain.Font = New-Object -TypeName System.Drawing.Font("MS Sans Serif",8)
    [void]$StatusStripMain.Items.Add($ToolStripStatusLabelMain)
    $FormMain.Controls.Add($StatusStripMain)

    [void]$FormMain.ShowDialog()
    $FormMain.Dispose()
}
#endregion

#region Handlers
$FormMain_Shown =
{
    $ToolStripStatusLabelMain.Text = "Ready"
    $StatusStripMain.Update()
    $FormMain.Activate()
}

$TextBoxName_TextChanged =
{
    if ($TextBoxName.TextLength -eq 0)
    {
        $ErrorProviderMain.Clear()
        $ButtonRun.Enabled = $false
    }
    elseif ($TextBoxName.Text -match "[^a-z0-9A-Z\-]")
    {
        $ErrorProviderMain.SetIconPadding($TextBoxName,-20)
        $ErrorProviderMain.SetError($TextBoxName,"The server name contains an invalid character.")
        $ButtonRun.Enabled = $false
    }
    else
    {
        $ErrorProviderMain.Clear()
        $ButtonRun.Enabled = $true
    }
}

$ButtonRun_Click = 
{
    $FormMain.Controls | Where-Object {$PSItem -isnot [System.Windows.Forms.StatusStrip]} | ForEach-Object {$PSItem.Enabled = $false}
    $FormMain.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    try
    {
        $ToolStripStatusLabelMain.Text = "Testing connection to remote server..."
        $StatusStripMain.Update()
        $ServerName = $TextBoxName.Text.Trim()
        $RestartTime = $TextBoxTime.Text.Trim()
        $RestartDate = $DateTimePickerDate.Text.Trim()
        if (Test-Connection $ServerName -quiet -count 1)
        {
            $ToolStripStatusLabelMain.Text = "Scheduling restart..."
            $StatusStripMain.Update()
            $Trigger = New-ScheduledTaskTrigger -Once -At "$RestartDate $RestartTime"
            $Trigger.EndBoundary = [datetime]::Parse("$RestartDate $RestartTime").AddMinutes(5).ToString('s')
            $Action = New-ScheduledTaskAction -Execute $TaskAction -Argument $TaskArgument
            $Settings = New-ScheduledTaskSettingsSet -DeleteExpiredTaskAfter 00:00:00
            Invoke-Command -ComputerName $ServerName -ScriptBlock {
                if (Get-ScheduledTask -TaskName $Using:TaskName -TaskPath "\$Using:TaskPath\" -ErrorAction SilentlyContinue)
                {
                    Unregister-ScheduledTask -TaskName $Using:TaskName -TaskPath "\$Using:TaskPath\" -Confirm:$false
                }
                Register-ScheduledTask -TaskName $Using:TaskName -Action $Using:Action -Trigger $Using:Trigger -TaskPath $Using:TaskPath -Settings $Using:Settings -User $Using:TaskUser
            }
            [void][System.Windows.Forms.MessageBox]::Show(
                "A one-time restart on $ServerName has been scheduled.`n`nDate:`t$RestartDate`nTime:`t$RestartTime",
                "Task Scheduled",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        else
        {
            [void][System.Windows.Forms.MessageBox]::Show(
                "The server $ServerName is not responding.`n`nPlease verify the name is correct and the server is online.",
                "Communication Problem",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    }
    catch
    {
        [void][System.Windows.Forms.MessageBox]::Show(
            $PSItem.Exception.Message + "`n`nPlease contact $SupportContact for technical support.",
            "Exception",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }

    $FormMain.Controls | ForEach-Object {$PSItem.Enabled = $true}
    $FormMain.ResetCursor()
    $ErrorProviderMain.Clear()
    $ButtonRun.Enabled = $false
    $TextBoxName.Clear()
    $TextBoxName.Focus()
    $ToolStripStatusLabelMain.Text = "Ready"
    $StatusStripMain.Update()
}
#endregion

#region Main
Invoke-Command -ScriptBlock $ShowFormMain
#endregion
