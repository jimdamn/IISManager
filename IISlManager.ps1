Add-Type -AssemblyName System.Windows.Forms

function LimitInput {
    if ($AppCode.Text.Length -gt 3 -or -not ($AppCode.Text -cmatch '^[a-zA-Z]{0,3}$')) {
        $AppCode.Text = $AppCode.Text.Substring(0, $AppCode.Text.Length - 1)
    }
}

function ClearSelectLists {

    $AppCode.Text = ""
    $slServerSelect.Items.Clear()
    $slAppPoolSelect.Items.Clear()
}

        
function ClearAPT {

    $lblErrMessages.Text = "Updating recycle time..."
    $lblErrMessages.Visible = $true
    $logFileSAP = "AppPoolModification.txt"

    if (!(Test-Path $logFileSAP)) {
        New-Item -ItemType File -Path $logFileSAP -Force
    }

    foreach ($server in $slServerSelect.SelectedItems) {
        foreach ($appPool in $slAppPoolSelect.SelectedItems) {
            try {

                $hour = $cbHour.SelectedItem
                $min = $cbMin.SelectedItem

                Invoke-Command -ComputerName $server -ScriptBlock {
                    param($appPool, $hour, $min)

                    Import-Module WebAdministration
                                      
                        Set-ItemProperty -Path "IIS:\AppPools\$appPool" -Name Recycling.periodicRestart.time -value 0.00:00:00
                        Clear-ItemProperty -Path "IIS:\AppPools\$appPool" -Name Recycling.periodicRestart.schedule
                        

                } -ArgumentList $appPool, $hour, $min

            } catch {
                $lblErrMessages.Text = "An error occurred while updating application pool '$appPool' on server '$server': $($_.Exception.Message)"
                $lblErrMessages.Visible = $true
            }
        }
    }

    $logMessageSAP = "$(Get-Date) - The application pool $($slAppPoolSelect.SelectedItems -join ",") has been set with no recycles on servers: $($slServerSelect.SelectedItems -join ",")"
    Add-Content -Path $logFileSAP -Value $logMessageSAP

    ClearSelectLists
    $lblErrMessages.Text = ""
    $lblErrMessages.Visible = $false
}

   function GetServers {
    $lblErrMessages.Text = "Please wait..."
    $lblErrMessages.Visible = $true 
    $rbEnv = if ($rbEnvTest.Checked) { "tv" } else { "pv" }
    $appCode = $AppCode.Text

    if ([string]::IsNullOrEmpty($appCode)) {
        $lblErrMessages.Text = "Please enter an Application Code."
        $lblErrMessages.Visible = $true
        return
    }

    try {
        $servers = Get-ADComputer -Filter "Name -like '$appCode*$rbEnv*'" | Where-Object {$_.Name -match "^$appCode.{6}$rbEnv.{2}$" -and (Get-WmiObject Win32_Service -Filter "Name='W3SVC'" -ComputerName $_.Name -ErrorAction SilentlyContinue)}
        $arrHostName = $servers.Name

        $slServerSelect.Items.Clear()
        foreach ($server in $arrHostName) {
            $slServerSelect.Items.Add($server)
        }
    } catch {
        $lblErrMessages.Text = "An error occurred while fetching servers: $($_.Exception.Message)"
        $lblErrMessages.Visible = $true
    }

    GetAppPools $arrHostName
}

function IISReset {
    Import-Module WebAdministration
    $lblErrMessages.Text = "Performing IIS Reset"
    $lblErrMessages.Visible = $true
    $logFile = "IISResetLog.txt"

    if (!(Test-Path $logFile)) {
        New-Item -ItemType File -Path $logFile -Force
    }

    $jobs = @()
    foreach ($server in $slServerSelect.SelectedItems) {
        try {
            
            Invoke-Command -ComputerName $server -ScriptBlock { Stop-WebSite -Name "HC" }
	        Start-Sleep -Seconds 15

            $job = Invoke-Command -ComputerName $server -ScriptBlock {
                param($server)
                iisreset /restart
                $iisRunning = $false
                while (!$iisRunning) {
                    Start-Sleep -Seconds 5
                    $iisRunning = Get-Service -Name "W3SVC" -ComputerName $server | Where-Object { $_.Status -eq "Running" }
                }
                $iisRunning
            } -ArgumentList $server -AsJob

            $websiteStatus = (Invoke-Command -ComputerName $server -ScriptBlock {Get-Website -Name 'HC'} | Select-Object -ExpandProperty State)
        	if ($websiteStatus -eq "Started") {
            	Start-Sleep -Seconds 15
        	} else {
            	Write-Host "Warning: HC website is not running!"
        	}

            $jobs += $job

        } catch {
            $lblErrMessages.Text = "An error occurred while performing IIS Reset on server '$server': $($_.Exception.Message)"
            $lblErrMessages.Visible = $true
        }
    }

    Wait-Job $jobs

    $failedJobs = $jobs | Where-Object { $_.State -eq "Failed" }
    if ($failedJobs) {
        $failedJob = $failedJobs[0]
        $errorMessage = "An error occurred while performing IIS Reset on server $($failedJob.PSComputerName): $($failedJob.ChildJobs[0].JobStateInfo.Reason)"
        $lblErrMessages.Text = $errorMessage
        $lblErrMessages.Visible = $true
    }

    $logMessage = "$(Get-Date) - Performed IIS Reset on servers: $($slServerSelect.SelectedItems -join ",")"
    Add-Content -Path $logFile -Value $logMessage

    ClearSelectLists
    $lblErrMessages.Text = ""
    $lblErrMessages.Visible = $false
}

function GetAppPools {
    $arrAP = @()
    foreach ($server in $arrHostName) {
        try {
            
            $appPools = Invoke-Command -ComputerName $server -ScriptBlock { 
            
            Import-Module WebAdministration
            Get-IISAppPool 
            }

            foreach ($appPool in $appPools) {
                $poolName = $appPool.Name
                if ($poolName -notin $arrAP) {
                    $arrAP += $poolName
                }
            }
        } catch {
            $lblErrMessages.Text = "An error occurred while fetching application pools from server {$server}: $($_.Exception.Message)"
            $lblErrMessages.Visible = $true
        }
    }

    
    $slAppPoolSelect.Items.Clear()
    foreach ($appPoolName in $arrAP) {
        $slAppPoolSelect.Items.Add($appPoolName)
    }
    $lblErrMessages.Text = ""
    $lblErrMessages.Visible = $false
}


function RecycleAppPools {

    $lblErrMessages.Text = "Performing App Pool recycle..."
    $lblErrMessages.Visible = $true

    $logFileAP = "AppPoolRecycleLog.txt"

    if (!(Test-Path $logFileAP)) {
        New-Item -ItemType File -Path $logFileAP -Force
    }

    foreach ($server in $slServerSelect.SelectedItems) {
        foreach ($appPool in $slAppPoolSelect.SelectedItems) {
            try {
                
                Invoke-Command -ComputerName $server -ScriptBlock {
                    param($appPool)

                    Import-Module WebAdministration
                    Restart-WebAppPool $appPool
                } -ArgumentList $appPool

                
                $logMessage = "$(Get-Date) - Recycled application pool '$appPool' on server '$server'"
                Add-Content -Path $logFile -Value $logMessage

            } catch {
                $lblErrMessages.Text = "An error occurred while recycling application pool '$appPool' on server '$server': $($_.Exception.Message)"
                $lblErrMessages.Visible = $true
            }
        }
    }

    $logMessageAP = "$(Get-Date) - Performed App Pool Recycle for $($slAppPoolSelect.SelectedItems -join ",") on servers: $($slServerSelect.SelectedItems -join ",")"
    Add-Content -Path $logFileAP -Value $logMessageAP
    
    $lblErrMessages.Text = ""
    $lblErrMessages.Visible = $false
    ClearSelectLists
}


$form = New-Object System.Windows.Forms.Form
$form.Text = "IIS Manager"
$form.Width = 325
$form.Height = 460

$lblACode = New-Object System.Windows.Forms.Label
$lblACode.Text = "Application Code:"
$lblACode.Location = New-Object System.Drawing.Point(20, 10)
$form.Controls.Add($lblACode)

$AppCode = New-Object System.Windows.Forms.TextBox
$AppCode.Location = New-Object System.Drawing.Point(20, 33)
$AppCode.Width = 100
$AppCode.MaxLength = 3
$AppCode.Add_TextChanged({ LimitInput })
$form.Controls.Add($AppCode)

$btnSubmit = New-Object System.Windows.Forms.Button
$btnSubmit.Text = "Get Servers"
$btnSubmit.Location = New-Object System.Drawing.Point(125, 31)
$btnSubmit.Add_Click({ GetServers })
$form.Controls.Add($btnSubmit)

$rbEnvTest = New-Object System.Windows.Forms.RadioButton
$rbEnvTest.Text = "Test"
$rbEnvTest.Location = New-Object System.Drawing.Point(20, 60)
$rbEnvTest.Checked = $true
$form.Controls.Add($rbEnvTest)

$rbEnvProd = New-Object System.Windows.Forms.RadioButton
$rbEnvProd.Text = "Production"
$rbEnvProd.Location = New-Object System.Drawing.Point(20, 80)
$form.Controls.Add($rbEnvProd)

$lblErrMessages = New-Object System.Windows.Forms.Label
$lblErrMessages.Location = New-Object System.Drawing.Point(20, 350)
$lblErrMessages.Width = 200
$lblErrMessages.Height = 100
$lblErrMessages.ForeColor = [System.Drawing.Color]::Red
$lblErrMessages.Visible = $false
$form.Controls.Add($lblErrMessages)

$slServerSelect = New-Object System.Windows.Forms.ListBox
$slServerSelect.Location = New-Object System.Drawing.Point(20, 110)
$slServerSelect.Width = 135
$slServerSelect.SelectionMode = "MultiExtended"
$form.Controls.Add($slServerSelect)


$slAppPoolSelect = New-Object System.Windows.Forms.ListBox
$slAppPoolSelect.Location = New-Object System.Drawing.Point(160, 110)
$slAppPoolSelect.Width = 135
$slAppPoolSelect.SelectionMode = "MultiExtended"
$form.Controls.Add($slAppPoolSelect)

$btnExRecycle = New-Object System.Windows.Forms.Button
$btnExRecycle.Text = "App Recycle"
$btnExRecycle.Location = New-Object System.Drawing.Point(20, 215)
$btnExRecycle.Width = 90
$btnExRecycle.Add_Click({ RecycleAppPools })
$form.Controls.Add($btnExRecycle)

$btnExClear = New-Object System.Windows.Forms.Button
$btnExClear.Text = "Clear"
$btnExClear.Location = New-Object System.Drawing.Point(210, 215)
$btnExClear.Width = 80
$btnExClear.Add_Click({ ClearSelectLists })
$form.Controls.Add($btnExClear)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "IIS Reset"
$btnReset.Location = New-Object System.Drawing.Point(115, 215)
$btnReset.Width = 90
$btnReset.Add_Click({ IISReset })
$form.Controls.Add($btnReset)

$lblSelectTimes                  = New-Object system.Windows.Forms.Label
$lblSelectTimes.text             = "Select time for daily recycle:"
$lblSelectTimes.AutoSize         = $true
$lblSelectTimes.width            = 25
$lblSelectTimes.height           = 10
$lblSelectTimes.location         = New-Object System.Drawing.Point(20,265)
$lblSelectTimes.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$form.Controls.Add($lblSelectTimes)

$cbHour                          = New-Object system.Windows.Forms.ComboBox
$cbHour.text                     = ""
$cbHour.width                    = 50
$cbHour.height                   = 30
@('01','02','03','04','05','06','07','08','09','10','11', '12', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23') | ForEach-Object {[void] $cbHour.Items.Add($_)}
$cbHour.SelectedIndex            = 4
$cbHour.location                 = New-Object System.Drawing.Point(20,282)
$cbHour.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$form.Controls.Add($cbHour)

$cbMin                           = New-Object system.Windows.Forms.ComboBox
$cbMin.text                      = ""
$cbMin.width                     = 50
$cbMin.height                    = 30
@('00','15','30','45') | ForEach-Object {[void] $cbMin.Items.Add($_)}
$cbMin.SelectedIndex            = 0
$cbMin.location                  = New-Object System.Drawing.Point(75,282)
$cbMin.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$form.Controls.Add($cbMin)

$btnSelectTime                   = New-Object system.Windows.Forms.Button
$btnSelectTime.text              = "Set Recycle Time"
$btnSelectTime.width             = 125
$btnSelectTime.height            = 30
$btnSelectTime.location          = New-Object System.Drawing.Point(20,312)
$btnSelectTime.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$btnSelectTime.Add_Click({ UpdAppPoolsConfig $cbHour.SelectedItem $cbMin.SelectedItem})
$form.Controls.Add($btnSelectTime)

$btnClearAPT = New-Object System.Windows.Forms.Button
$btnClearAPT.Text = "No Recycles"
$btnClearAPT.Location = New-Object System.Drawing.Point(150, 312)
$btnClearAPT.Width = 90
$btnClearAPT.height            = 30
$btnClearAPT.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$btnClearAPT.Add_Click({ ClearAPT })
$form.Controls.Add($btnClearAPT)     

$form.ShowDialog()
