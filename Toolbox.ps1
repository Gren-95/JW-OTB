Write-Output " _____   __      __      _____    ______  ____      ";
Write-Output "/\___ \ /\ \  __/\ \    /\  __\`\ /\__  _\/\  _\`\    ";
Write-Output "\/__/\ \\ \ \/\ \ \ \   \ \ \/\ \\/_/\ \/\ \ \L\ \  ";
Write-Output "   _\ \ \\ \ \ \ \ \ \   \ \ \ \ \  \ \ \ \ \  _ <' ";
Write-Output "  /\ \_\ \\ \ \_/ \_\ \   \ \ \_\ \  \ \ \ \ \ \L\ \";
Write-Output "  \ \____/ \ \`\___x___/    \ \_____\  \ \_\ \ \____/";
Write-Output "   \/___/   '\/__//__/      \/_____/   \/_/  \/___/ ";
Write-Output "                                                    ";
Write-Output "                                                    ";


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$ApplicationForm = New-Object System.Windows.Forms.Form
$ApplicationForm.StartPosition = "CenterScreen"
$ApplicationForm.Topmost = $false 
$ApplicationForm.Size = "700,500"
$ApplicationForm.FormBorderStyle = 'Fixed3D'
$ApplicationForm.MinimizeBox = $false
$ApplicationForm.MaximizeBox = $false
$ApplicationForm.Text = "JW Open ToolBox"
#$ApplicationForm.Topmost = $true

# tab Control Window
$FormTabControl = New-object System.Windows.Forms.TabControl 
$FormTabControl.Size = "681,500" 
$FormTabControl.Location = "0,0" 
$ApplicationForm.Controls.Add($FormTabControl)


# TAB 1
#----------------------------------------------


$Tab1_row1_distance=25
$Tab1_row2_distance=235
$Tab1_row3_distance=445

$Tab1 = New-object System.Windows.Forms.Tabpage
$Tab1.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab1.Name = "Tab1" 
$Tab1.Text = "Configure” 
$FormTabControl.Controls.Add($Tab1)

# Config label 1
$Tab1_label1 = New-Object System.Windows.Forms.Label
$Tab1_label1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,10)
$Tab1_label1.Font = New-Object System.Drawing.Font('verdana',16)
$Tab1_label1.AutoSize = $true
$Tab1_label1.ForeColor = "#000000"
$Tab1_label1.Text = ("Config")
$Tab1.Controls.Add($Tab1_label1)

# Update label 2
$Tab1_label2 = New-Object System.Windows.Forms.Label
$Tab1_label2.Location = New-Object System.Drawing.Point($Tab1_row2_distance,10)
$Tab1_label2.Font = New-Object System.Drawing.Font('verdana',16)
$Tab1_label2.AutoSize = $true
$Tab1_label2.ForeColor = "#000000"
$Tab1_label2.Text = ("Update")
$Tab1.Controls.Add($Tab1_label2)

# Domian label 3
$Tab1_label3 = New-Object System.Windows.Forms.Label
$Tab1_label3.Location = New-Object System.Drawing.Point($Tab1_row3_distance,10)
$Tab1_label3.Font = New-Object System.Drawing.Font('verdana',16)
$Tab1_label3.AutoSize = $true
$Tab1_label3.ForeColor = "#000000"
$Tab1_label3.Text = ("Domian")
$Tab1.Controls.Add($Tab1_label3)


# Printers label 4
$Tab1_label4 = New-Object System.Windows.Forms.Label
$Tab1_label4.Location = New-Object System.Drawing.Point($Tab1_row2_distance,175)
$Tab1_label4.Font = New-Object System.Drawing.Font('verdana',16)
$Tab1_label4.AutoSize = $true
$Tab1_label4.ForeColor = "#000000"
$Tab1_label4.Text = ("Printers")
$Tab1.Controls.Add($Tab1_label4)

# Printers label 4
$Tab1_label4 = New-Object System.Windows.Forms.Label
$Tab1_label4.Location = New-Object System.Drawing.Point($Tab1_row3_distance,175)
$Tab1_label4.Font = New-Object System.Drawing.Font('verdana',16)
$Tab1_label4.AutoSize = $true
$Tab1_label4.ForeColor = "#000000"
$Tab1_label4.Text = ("Wi-Fi")
$Tab1.Controls.Add($Tab1_label4)

# Lots of buttons



$Tab1_appbutton1 = New-Object System.Windows.Forms.Button
$Tab1_appbutton1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,50)
$Tab1_appbutton1.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton1.Text = "Device Manager - done"
$Tab1_appbutton1.Add_Click({Start-Process hdwwiz.cpl})
$tab1.Controls.Add($Tab1_appbutton1)

$Tab1_appbutton2 = New-Object System.Windows.Forms.Button
$Tab1_appbutton2.Location = New-Object System.Drawing.Point($Tab1_row1_distance,85)
$Tab1_appbutton2.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton2.Text = "Display Properties - done"
$Tab1_appbutton2.Add_Click({Start-Process desk.cpl})
$tab1.Controls.Add($Tab1_appbutton2)

$Tab1_appbutton3 = New-Object System.Windows.Forms.Button
$Tab1_appbutton3.Location = New-Object System.Drawing.Point($Tab1_row1_distance,120)
$Tab1_appbutton3.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton3.Text = "Control Panel - done"
$Tab1_appbutton3.Add_Click({Start-Process control})
$tab1.Controls.Add($Tab1_appbutton3)

$Tab1_appbutton4 = New-Object System.Windows.Forms.Button
$Tab1_appbutton4.Location = New-Object System.Drawing.Point($Tab1_row1_distance,155)
$Tab1_appbutton4.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton4.Text = "Sound and Audio - done"
$Tab1_appbutton4.Add_Click({Start-Process mmsys.cpl})
$tab1.Controls.Add($Tab1_appbutton4)

$Tab1_appbutton5 = New-Object System.Windows.Forms.Button
$Tab1_appbutton5.Location = New-Object System.Drawing.Point($Tab1_row1_distance,190)
$Tab1_appbutton5.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton5.Text = "Network Properties - done"
$Tab1_appbutton5.Add_Click({Start-Process ncpa.cpl})
$tab1.Controls.Add($Tab1_appbutton5)

$Tab1_appbutton6 = New-Object System.Windows.Forms.Button
$Tab1_appbutton6.Location = New-Object System.Drawing.Point($Tab1_row1_distance,225)
$Tab1_appbutton6.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton6.Text = "Settings - done"
$Tab1_appbutton6.Add_Click({Start-Process ms-settings:})
$tab1.Controls.Add($Tab1_appbutton6)


$Tab1_appbutton9 = New-Object System.Windows.Forms.Button
$Tab1_appbutton9.Location = New-Object System.Drawing.Point($Tab1_row1_distance,260)
$Tab1_appbutton9.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton9.Text = "Calibrate touchscreen (Admin) - done"
$Tab1_appbutton9.Add_Click({Start-Process tabcal.exe})
$tab1.Controls.Add($Tab1_appbutton9)


$Tab1_appbutton10 = New-Object System.Windows.Forms.Button
$Tab1_appbutton10.Location = New-Object System.Drawing.Point($Tab1_row1_distance,295)
$Tab1_appbutton10.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton10.Text = "Add/Remove Applications - done"
$Tab1_appbutton10.Add_Click({Start-Process appwiz.cpl})
$tab1.Controls.Add($Tab1_appbutton10)


$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row1_distance,330)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "IT NAS - done"
$Tab1_appbutton11.Add_Click({explorer \\10.72.7.1\IT})
$tab1.Controls.Add($Tab1_appbutton11)







$Tab1_appbutton111 = New-Object System.Windows.Forms.Button
$Tab1_appbutton111.Location = New-Object System.Drawing.Point($Tab1_row2_distance,50)
$Tab1_appbutton111.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton111.Text = "Get Winget"
$Tab1_appbutton111.Add_Click({
	$URL = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
	$URL = (Invoke-WebRequest -Uri $URL).Content | ConvertFrom-Json |
			Select-Object -ExpandProperty "assets" |
			Where-Object "browser_download_url" -Match '.msixbundle' |
			Select-Object -ExpandProperty "browser_download_url"
# download
	Invoke-WebRequest -Uri $URL -OutFile "Setup.msix" -UseBasicParsing

# install
	Add-AppxPackage -Path "Setup.msix"

# delete file
	Remove-Item "Setup.msix"
})
$tab1.Controls.Add($Tab1_appbutton111)

$Tab1_appbutton112 = New-Object System.Windows.Forms.Button
$Tab1_appbutton112.Location = New-Object System.Drawing.Point($Tab1_row2_distance,85)
$Tab1_appbutton112.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton112.Text = "Install winget updates (Apps) - done - admin"
$Tab1_appbutton112.Add_Click({start-process powershell -verb runas {
		winget update --verbose 
		winget upgrade --all --accept-source-agreements --accept-source-agreements 
	}
	

})
$tab1.Controls.Add($Tab1_appbutton112)

$Tab1_appbutton113 = New-Object System.Windows.Forms.Button
$Tab1_appbutton113.Location = New-Object System.Drawing.Point($Tab1_row2_distance,120)
$Tab1_appbutton113.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton113.Text = "Update Office - done"
$Tab1_appbutton113.Add_Click({
	#info text
Write-Host ""
Write-Host "You can download this script at : " -NoNewline
Write-Host "https://aka.ms/update365" -ForegroundColor Yellow
Write-Host ""
Write-Host "Update history for Office 365 ProPlus : "-NoNewline
Write-Host "https://docs.microsoft.com/en-us/officeupdates/update-history-office365-proplus-by-date" -ForegroundColor Yellow
Write-Host "Release information for updates to Office 365 ProPlus : " -NoNewline
Write-Host "https://docs.microsoft.com/en-us/officeupdates/release-notes-office365-proplus" -ForegroundColor Yellow
Write-Host "Update history for Office Insider for Windows desktop: " -NoNewline
Write-Host "https://support.office.com/en-us/article/update-history-for-office-insider-for-windows-desktop-64bbb317-972a-4933-8b82-cc866f0b067c" -ForegroundColor Yellow
Write-Host ""


if (Test-Path "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe"){
$ErrorActionPreference= 'silentlycontinue'

Write-Host "Getting a release information of Office 365......."

#get release version
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$URI="https://docs.microsoft.com/en-us/officeupdates/update-history-office365-proplus-by-date"
$HTML = Invoke-WebRequest -Uri $URI
$result = $HTML.Content


#insider channel data
$InsiderfastURL="https://docs.microsoft.com/en-us/officeupdates/update-history-beta-channel"
$InsiderfastHTML = Invoke-WebRequest -Uri $InsiderfastURL
$InsiderfastResult = [regex]::matches( $InsiderfastHTML.Content , '(January|February|March|Arpil|May|June|July|August|September|October|November|December)\s\d{1,2},\s\d{4}</strong><br/>\s*Version\s\d{4}\s\(Build\s\d{4,5}\.\d{4,5}\)')

$InsiderslowURL="https://docs.microsoft.com/en-us/officeupdates/update-history-current-channel-preview"
$InsiderslowHTML = Invoke-WebRequest -Uri $InsiderslowURL
$InsiderslowResult = [regex]::matches( $InsiderslowHTML.Content , '(January|February|March|Arpil|May|June|July|August|September|October|November|December)\s\d{1,2},\s\d{4}</strong><br/>\s*Version\s\d{4}\s\(Build\s\d{4,5}\.\d{4,5}\)')

$InsiderFastBuild = [regex]::Matches($InsiderfastResult,"Version \d{4} \(Build \d{4,5}\.\d{4,5}\)")
$InsiderSlowBuild = [regex]::Matches($InsiderslowResult,"Version \d{4} \(Build \d{4,5}\.\d{4,5}\)")
$InsiderFastDate = [regex]::Matches($InsiderfastResult,"(January|February|March|Arpil|May|June|July|August|September|October|November|December)\s\d{1,2},\s\d{4}")
$InsiderSlowDate = [regex]::Matches($InsiderslowResult,"(January|February|March|Arpil|May|June|July|August|September|October|November|December)\s\d{1,2},\s\d{4}")


#Channel Name
$nameCurrent = "Currnet Channel (Monthly Channel)"
$nameDeferred = "Semi-Annual Enterprise Channel (Semi-Annual Channel)"
$nameFirstdeferred = "Semi-Annual Enterprise Channel (Preview) (Semi-Annual Channel (Targeted))"
$nameInsiderfast = "Beta Channel (Insider Fast)"
$nameInsiderslow = "Current Channel (Preview) (Insider Slow)"
$nameDevMain = "DevMain Channel (Dogfood)"
$nameMonthlyEnt = "Montly Enterprise Channel"

#CDNBaseUrl
$CDNBaseUrlCurrent = "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
$CDNBaseUrlDeferred = "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
$CDNBaseUrlFirstDeferred = "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf"
$CDNBaseUrlInsiderFast= "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f"
$CDNBaseUrlInsiderSlow= "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be"
$CDNBaseUrlDevMain = "http://officecdn.microsoft.com/pr/ea4a4090-de26-49d7-93c1-91bff9e53fc3"
$CDNBaseUrlMonthlyEnt = "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6"

#Channel filter
$current = [regex]::matches( $result, '<a href=\"(monthly-channel|current-channel)(.*?)</a>')
$deferred = [regex]::matches( $result, '<a href="(semi-annual-channel-(\d{4})|semi-annual-enterprise-channel#)(.*?)</a>') #Thanks for correction by tobiasabele, https://github.com/tobiasabele
$firstDeferred = [regex]::matches( $result, '<a href=\"(semi-annual-channel-targeted-(\d{4})|semi-annual-enterprise-channel-preview)(.*?)</a>')
$monthlyEnt = [regex]::Matches($result, '<a href=\"monthly-enterprise-channel(.*?)</a>')
$ChannelChanged = $false


#Form 

Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Office 365 Update Tool"
$Form.Size = New-Object System.Drawing.Size(370,120)
$form.MaximumSize = New-Object System.Drawing.Size(370,150)
$Form.MinimumSize = New-Object System.Drawing.Size(370,150)
$CenterScreen = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$Form.StartPosition = $CenterScreen
$Form.TopMost = $True

$CmbChannel = New-Object System.Windows.Forms.ComboBox
$CmbChannel.Text = "Select release channel..."
$CmbChannel.Location = New-Object System.Drawing.Point(10,15)
$CmbChannel.Size = New-Object System.Drawing.Size(330,80)
$Form.controls.Add($CmbChannel)

$CmbBuild = New-Object System.Windows.Forms.ComboBox
$CmbBuild.Text = "Select release channel first..."
$CmbBuild.Location = New-Object System.Drawing.Point(10,40)
$CmbBuild.Size = New-Object System.Drawing.Size(330,400)
$Form.Controls.Add($CmbBuild)

$BtnUpdate = New-Object System.Windows.Forms.Button
$BtnUpdate.Text = "Update"
$BtnUpdate.Location = New-Object System.Drawing.Point(265,75)
$BtnUpdate.Enabled = $false
$Form.Controls.Add($BtnUpdate)

$ChkUpdate = New-Object System.Windows.Forms.Checkbox
$ChkUpdate.Text = "Disable updates"
$ChkUpdate.Location = New-Object System.Drawing.Point(10,75)
$ChkUpdate.Size = New-Object System.Drawing.Size(200,20)
$Form.Controls.Add($ChkUpdate)


#Cmb contents
$CmbChannel.Items.Add($nameCurrent) >> $null
$CmbChannel.Items.Add($nameMonthlyEnt) >> $null
$CmbChannel.Items.Add($nameDeferred) >> $null
$CmbChannel.Items.Add($nameFirstdeferred) >> $null
$CmbChannel.Items.Add($nameInsiderfast) >> $null
$CmbChannel.Items.Add($nameInsiderslow) >> $null
$CmbChannel.Items.Add($nameDevMain) >> $null

#Event handler

$CmbChannel_SelectedIndexChanged =
{
    $BtnUpdate.Enabled = $false
    if($CmbChannel.Text -eq $nameCurrent){
        $CmbBuild.Items.Clear()
        $CmbBuild.Text = "Select build number..."

        for($i=0;$i -lt $current.count;$i++){
            $date_build = ([regex]::matches($current.value[$i],'Version \d{4} \(Build \d{4,5}\.\d{4,5}\)' )).value
            $CmbBuild.Items.Add($date_build)
        }
    }
    elseif($CmbChannel.Text -eq $nameMonthlyEnt){
        $CmbBuild.Items.Clear()
        $CmbBuild.Text = "Select build number..."
        for($i=0;$i -lt $monthlyEnt.count;$i++){
            $date_build = ([regex]::matches($monthlyEnt.value[$i],'Version \d{4} \(Build \d{4,5}\.\d{4,5}\)' )).value
            $CmbBuild.Items.Add($date_build)
        }


    }
    elseif($CmbChannel.Text -eq $nameDeferred){
        $CmbBuild.Items.Clear()
        $CmbBuild.Text = "Select build number..."
        for($i=0;$i -lt $deferred.count;$i++){
            $date_build = ([regex]::matches($deferred.value[$i],'Version \d{4} \(Build \d{4,5}\.\d{4,5}\)' )).value
            $CmbBuild.Items.Add($date_build)
        }


    }
    elseif($CmbChannel.Text -eq $nameFirstdeferred){
        $CmbBuild.Items.Clear()
        $CmbBuild.Text = "Select build number..."    
        for($i=0;$i -lt $firstDeferred.count;$i++){
            $date_build = ([regex]::matches($firstDeferred.value[$i],'Version \d{4} \(Build \d{4,5}\.\d{4,5}\)' )).value
            $CmbBuild.Items.Add($date_build)
        }    
   
    }

    elseif($CmbChannel.Text -eq $nameInsiderfast){
        $CmbBuild.Items.Clear()
        $CmbBuild.Text = "Select build number..."    
        for($i=0;$i -lt $InsiderFastBuild.count;$i++){
            $CmbBuild.Items.Add($InsiderFastBuild[$i].value + " " +$InsiderFastDate[$i].value)
        }    
   
    }

    elseif($CmbChannel.Text -eq $nameInsiderslow){
        $CmbBuild.Items.Clear()
        $CmbBuild.Text = "Select build number..."    
        for($i=0;$i -lt $InsiderSlowBuild.count;$i++){
            $CmbBuild.Items.Add($InsiderSlowBuild[$i].value + " " +$InsiderslowDate[$i].value)
        }    
   
    }

    elseif($CmbChannel.Text -eq $nameDevMain){
        $BtnUpdate.Enabled = $true
        $CmbBuild.Items.Clear()
        $CmbBuild.Text = "Update to the latest build."       
    }


}

$CmbBuild_SelectedIndexChanged = { $BtnUpdate.Enabled = $True }

$BtnUpdate_Click =
{
    if($CmbChannel.Text -eq $nameCurrent){

        if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlCurrent)        {
            $ChannelChanged = $true
            Start-Process powershell.exe -Verb runAs{
            Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
            Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
            }
        }
    }
    
    elseif($CmbChannel.Text -eq $nameDeferred){
        if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlDeferred){         
            $ChannelChanged = $true
            Start-Process powershell.exe -Verb runAs{
            Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
            Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
            }
        }
    }

    elseif($CmbChannel.Text -eq $nameMonthlyEnt){
        if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlMonthlyEnt){         
            $ChannelChanged = $true
            Start-Process powershell.exe -Verb runAs{
            Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6"
            Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
            }
        }
    }

    elseif($CmbChannel.Text -eq $nameFirstdeferred){
        if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlFirstDeferred){         
            $ChannelChanged = $true
            Start-Process powershell.exe -Verb runAs{
            Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf"
            Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
            }
        }
    }
    elseif($CmbChannel.Text -eq $nameInsiderfast){
        if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlInsiderFast){         
            $ChannelChanged = $true
            Start-Process powershell.exe -Verb runAs{
            Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f"
            Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
            }
        }
    }

    elseif($CmbChannel.Text -eq $nameInsiderslow){
        if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlInsiderSlow){         
            $ChannelChanged = $true
            Start-Process powershell.exe -Verb runAs{
            Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be"
            Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
            }
        }
    }

    elseif($CmbChannel.Text -eq $nameDevMain){
        
        if((Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration).CDNBaseUrl -ne $CDNBaseUrlDevMain){         
            $ChannelChanged = $true
            Start-Process powershell.exe -Verb runAs{
            Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name CDNBaseUrl -Value "http://officecdn.microsoft.com/pr/ea4a4090-de26-49d7-93c1-91bff9e53fc3"
            Remove-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Updates -Name UpdateToVersion
            }
        }
    }

    $Form.Close()

        if($ChannelChanged -eq $true){
            Write-Host "Modifing registry value for "$CmbChannel.text -NoNewline
            for($c = 0; $c -lt 5; $c++){
                Write-Host "." -NoNewline
                Start-Sleep -Seconds 1 
            }
            Write-Host ""
        }

    
    
    if($CmbChannel.Text -ne $nameDevMain){
        $build = "16.0."+(($CmbBuild.text -split "Build ")[1] -split "\)")[0]
        & "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user updatetoversion=$build}
    
    else{& "$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user}

    Write-Host "Updating Office 365......."

    if($ChkUpdate.Checked -eq $true){
        Write-Host "Disable updates......."
        Start-Process powershell.exe -Verb runAs{
        Set-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name UpdatesEnabled -Value "False"
        }

    }

}

$CmbChannel.add_SelectedIndexChanged($CmbChannel_SelectedIndexChanged)
$BtnUpdate.add_Click($BtnUpdate_Click)
$CmbBuild.add_SelectedIndexChanged($CmbBuild_SelectedIndexChanged)
$Form.ShowDialog()
}

else {
    Write-Host "Please verify Office 365 is installed correctly. Can't find '$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe'" -ForegroundColor Yellow
}
})
$tab1.Controls.Add($Tab1_appbutton113)

$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row2_distance,155)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Dell Command | Update - WIP"
$Tab1_appbutton11.Add_Click({Start-Process 
	#'C:\Program Files\Dell\CommandUpdate\dcu-cli.exe'

	# get latest download url





})
$tab1.Controls.Add($Tab1_appbutton11)




$Tab1_appbutton7 = New-Object System.Windows.Forms.Button
$Tab1_appbutton7.Location = New-Object System.Drawing.Point($Tab1_row2_distance,225)
$Tab1_appbutton7.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton7.Text = "Printers - Done"
$Tab1_appbutton7.Add_Click({Start-Process control printers})
$tab1.Controls.Add($Tab1_appbutton7)


$Tab1_appbutton8 = New-Object System.Windows.Forms.Button
$Tab1_appbutton8.Location = New-Object System.Drawing.Point($Tab1_row2_distance,260)
$Tab1_appbutton8.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton8.Text = "EERAKSRV02 - done"
$Tab1_appbutton8.Add_Click({explorer \\eeraksrv02})
$tab1.Controls.Add($Tab1_appbutton8)




$Tab1_domainbutton1 = New-Object System.Windows.Forms.Button
$Tab1_domainbutton1.Location = New-Object System.Drawing.Point($Tab1_row3_distance,50)
$Tab1_domainbutton1.Size = New-Object System.Drawing.Size(200,35)
$Tab1_domainbutton1.Text = "Add to Domain - WIP"
$Tab1_appbutton11.Add_Click({
	Add-Computer -DomainName jwinc.jeld-wen.com
})
$tab1.Controls.Add($Tab1_domainbutton1)

$Tab1_domainbutton2 = New-Object System.Windows.Forms.Button
$Tab1_domainbutton2.Location = New-Object System.Drawing.Point($Tab1_row3_distance,85)
$Tab1_domainbutton2.Size = New-Object System.Drawing.Size(200,35)
$Tab1_domainbutton2.Text = "Domain Panel - done"
$Tab1_domainbutton2.Add_Click({Start-Process sysdm.cpl})
$tab1.Controls.Add($Tab1_domainbutton2)

$Tab1_domainbutton3 = New-Object System.Windows.Forms.Button
$Tab1_domainbutton3.Location = New-Object System.Drawing.Point($Tab1_row3_distance,120)
$Tab1_domainbutton3.Size = New-Object System.Drawing.Size(200,35)
$Tab1_domainbutton3.Text = "WIP"
$Tab1_domainbutton3.Add_Click({Start-Process desk.cpl})
$tab1.Controls.Add($Tab1_domainbutton3)



$Tab1_wifi1 = New-Object System.Windows.Forms.Button
$Tab1_wifi1.Location = New-Object System.Drawing.Point($Tab1_row3_distance,225)
$Tab1_wifi1.Size = New-Object System.Drawing.Size(200,35)
$Tab1_wifi1.Text = "JWWLAN - done"
$Tab1_wifi1.Add_Click({
    netsh wlan connect ssid="JWWLAN" name="JWWLAN"
})
$tab1.Controls.Add($Tab1_wifi1)

$Tab1_wifi2 = New-Object System.Windows.Forms.Button
$Tab1_wifi2.Location = New-Object System.Drawing.Point($Tab1_row3_distance,260)
$Tab1_wifi2.Size = New-Object System.Drawing.Size(200,35)
$Tab1_wifi2.Text = "Guest"
$Tab1_wifi2.Add_Click({
	netsh wlan connect ssid="Külalised - Visitors" name="Kõlalised - Visitors"
})
$tab1.Controls.Add($Tab1_wifi2)

$Tab1_wifi3 = New-Object System.Windows.Forms.Button
$Tab1_wifi3.Location = New-Object System.Drawing.Point($Tab1_row3_distance,295)
$Tab1_wifi3.Size = New-Object System.Drawing.Size(200,35)
$Tab1_wifi3.Text = "JW Primary"
$Tab1_wifi3.Add_Click({
	netsh wlan connect ssid="JW Primary" name="JW Primary"
})
$tab1.Controls.Add($Tab1_wifi3)
























# TAB 2
#----------------------------------------------


$Tab2 = New-object System.Windows.Forms.Tabpage
$Tab2.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab2.Name = "Tab2" 
$Tab2.Text = "Install” 
$FormTabControl.Controls.Add($Tab2)

# label
$Tab2_label1 = New-Object System.Windows.Forms.Label
$Tab2_label1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,10)
$Tab2_label1.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label1.AutoSize = $true
$Tab2_label1.ForeColor = "#000000"
$Tab2_label1.Text = ("Install")
$Tab2.Controls.Add($Tab2_label1)


# label
$Tab2_label3 = New-Object System.Windows.Forms.Label
$Tab2_label3.Location = New-Object System.Drawing.Point($Tab1_row3_distance,10)
$Tab2_label3.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label3.AutoSize = $true
$Tab2_label3.ForeColor = "#000000"
$Tab2_label3.Text = ("AppX")
$Tab2.Controls.Add($Tab2_label3)












$Tab2_installbutton1 = New-Object System.Windows.Forms.Button
$Tab2_installbutton1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,50)
$Tab2_installbutton1.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton1.Text = "Digidoc4 - to be tested"
$Tab2_installbutton1.Add_Click({
	#Invoke-WebRequest -Uri https://github.com/open-eid/DigiDoc4-Client/releases/download/v4.3.0/Digidoc4_Client-4.3.0.4438.x64.et-EE.qt.msi -OutFile "C:\Users\Public\Documents\digidoc.msi"
	#Write-Host "downloaded"
	#C:\Users\Public\Documents\digidoc.msi /quiet
	#Write-Host "installed"

	$URL = "https://api.github.com/repos/open-eid/DigiDoc4-Client/releases/latest"
	$URL = (Invoke-WebRequest -Uri $URL).Content | ConvertFrom-Json |
			Select-Object -ExpandProperty "assets" |
			Where-Object "browser_download_url" -Match 'EE.qt.msi' |
			Select-Object -ExpandProperty "browser_download_url"
# download
	Invoke-WebRequest -Uri $URL -OutFile "C:\Users\Public\Documents\digidoc.msi" -UseBasicParsing

# install
C:\Users\Public\Documents\digidoc.msi /quiet

# delete file
#	Remove-Item "digidoc.msi"


})
$tab2.Controls.Add($Tab2_installbutton1)

$Tab2_installbutton2 = New-Object System.Windows.Forms.Button
$Tab2_installbutton2.Location = New-Object System.Drawing.Point($Tab1_row1_distance,85)
$Tab2_installbutton2.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton2.Text = "Barcode Fonts - to be tested"
$Tab2_installbutton2.Add_Click({
    xcopy "\\10.72.7.1\IT\FRE3OF9X.TTF" "%userprofile%\AppData\Local\Microsoft\Windows\Fonts"
    xcopy "\\10.72.7.1\IT\FREE3OF9.TTF" "%userprofile%\AppData\Local\Microsoft\Windows\Fonts"
})
$tab2.Controls.Add($Tab2_installbutton2)

$Tab2_installbutton3 = New-Object System.Windows.Forms.Button
$Tab2_installbutton3.Location = New-Object System.Drawing.Point($Tab1_row1_distance,120)
$Tab2_installbutton3.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton3.Text = "Java - to be tested"
$Tab2_installbutton3.Add_Click({Start-Process "\\10.72.7.1\it\JAVA jre-8u221-windows-i586.exe"})
$tab2.Controls.Add($Tab2_installbutton3)

$Tab2_installbutton4 = New-Object System.Windows.Forms.Button
$Tab2_installbutton4.Location = New-Object System.Drawing.Point($Tab1_row1_distance,155)
$Tab2_installbutton4.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton4.Text = "Cisco Anyconnect - to be tested"
$Tab2_installbutton4.Add_Click({Start-Process "\\10.72.7.1\it\anyconnect-win-4.10.05095-core-vpn-webdeploy-k9.msi"})
$tab2.Controls.Add($Tab2_installbutton4)

$Tab2_installbutton5 = New-Object System.Windows.Forms.Button
$Tab2_installbutton5.Location = New-Object System.Drawing.Point($Tab1_row1_distance,190)
$Tab2_installbutton5.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton5.Text = "Jabra Connect - to be tested"
$Tab2_installbutton5.Add_Click({Start-Process "\\10.72.7.1\it\JabraDirectSetup.exe"})
$tab2.Controls.Add($Tab2_installbutton5)

$Tab2_installbutton6 = New-Object System.Windows.Forms.Button
$Tab2_installbutton6.Location = New-Object System.Drawing.Point($Tab1_row1_distance,225)
$Tab2_installbutton6.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton6.Text = "Adobe Acrobat - to be tested"
$Tab2_installbutton6.Add_Click({Start-Process "\\10.72.7.1\it\!ajutine\Efe\1 - Installers\Adobe Acrobat Reader DC\readerdc64_en_l_cra_mdr_install.exe"})
$tab2.Controls.Add($Tab2_installbutton6)


$Tab2_installbutton7 = New-Object System.Windows.Forms.Button
$Tab2_installbutton7.Location = New-Object System.Drawing.Point($Tab1_row1_distance,260)
$Tab2_installbutton7.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton7.Text = "Add Active Directory Users and Computers - to be tested"
$Tab2_installbutton7.Add_Click({
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
	Restart-Service wuauserv
    Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
})
$tab2.Controls.Add($Tab2_installbutton7)


$Tab2_installbutton8 = New-Object System.Windows.Forms.Button
$Tab2_installbutton8.Location = New-Object System.Drawing.Point($Tab1_row1_distance,295)
$Tab2_installbutton8.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton8.Text = "MS Office - to be tested"
$Tab2_installbutton8.Add_Click({Start-Process "\\10.72.7.1\it\!ajutine\Efe\1 - Installers\Office new\OfficeSetup.exe"})
$tab2.Controls.Add($Tab2_installbutton8)


$Tab2_installbutton9 = New-Object System.Windows.Forms.Button
$Tab2_installbutton9.Location = New-Object System.Drawing.Point($Tab1_row1_distance,330)
$Tab2_installbutton9.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton9.Text = "Firefox - to be tested"
$Tab2_installbutton9.Add_Click({Start-Process "\\10.72.7.1\it\!ajutine\Efe\1 - Installers\Firefox\Firefox Installer.exe"})
$tab2.Controls.Add($Tab2_installbutton9)


$Tab2_installbutton10 = New-Object System.Windows.Forms.Button
$Tab2_installbutton10.Location = New-Object System.Drawing.Point($Tab1_row1_distance,365)
$Tab2_installbutton10.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton10.Text = "Powertoys - WIP"
$Tab2_installbutton10.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton10)


$Tab2_installbutton11 = New-Object System.Windows.Forms.Button
$Tab2_installbutton11.Location = New-Object System.Drawing.Point($Tab1_row1_distance,400)
$Tab2_installbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton11.Text = "WSL - to be tested"
$Tab2_installbutton11.Add_Click({start-process powershell -verb runas {
	Enable-WindowsOptionalFeature -Online -FeatureName "VirtualMachinePlatform" -All
	Enable-WindowsOptionalFeature -Online -FeatureName "Microsoft-Windows-Subsystem-Linux" -All
	Write-Host "WSL is now installed and configured. Please Reboot before using."
	}})
$tab2.Controls.Add($Tab2_installbutton11)


$Tab2_installbutton12 = New-Object System.Windows.Forms.Button
$Tab2_installbutton12.Location = New-Object System.Drawing.Point($Tab1_row2_distance,50)
$Tab2_installbutton12.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton12.Text = "PMS - WIP"
$Tab2_installbutton12.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton12)

$Tab2_installbutton13 = New-Object System.Windows.Forms.Button
$Tab2_installbutton13.Location = New-Object System.Drawing.Point($Tab1_row2_distance,85)
$Tab2_installbutton13.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton13.Text = "TITAN - WIP"
$Tab2_installbutton13.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton13)

$Tab2_installbutton14 = New-Object System.Windows.Forms.Button
$Tab2_installbutton14.Location = New-Object System.Drawing.Point($Tab1_row2_distance,120)
$Tab2_installbutton14.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton14.Text = "PPS Shortcut - WIP"
$Tab2_installbutton14.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton14)


$Tab2_installbutton15 = New-Object System.Windows.Forms.Button
$Tab2_installbutton15.Location = New-Object System.Drawing.Point($Tab1_row2_distance,155)
$Tab2_installbutton15.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton15.Text = "PPS Edge fix"
$Tab2_installbutton15.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton15)

$Tab2_installbutton15 = New-Object System.Windows.Forms.Button
$Tab2_installbutton15.Location = New-Object System.Drawing.Point($Tab1_row2_distance,190)
$Tab2_installbutton15.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton15.Text = "SAP - WIP"
$Tab2_installbutton15.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton15)

$Tab2_installbutton15 = New-Object System.Windows.Forms.Button
$Tab2_installbutton15.Location = New-Object System.Drawing.Point($Tab1_row2_distance,225)
$Tab2_installbutton15.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton15.Text = "SAP2  - WIP"
$Tab2_installbutton15.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton15)

















$Tab2_pkgmgrbutton4 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton4.Location = New-Object System.Drawing.Point($Tab1_row3_distance,50)
$Tab2_pkgmgrbutton4.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton4.Text = "Reinstall MS Store"
$Tab2_pkgmgrbutton4.Add_Click({start-process powershell -verb runas {
Get-AppxPackage -allusers Microsoft.WindowsStore | ForEach-Object {Add-AppxPackage -DisableDevelopmentMode -Register “$($_.InstallLocation)\AppXManifest.xml”}
}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton4)

$Tab2_pkgmgrbutton5 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton5.Location = New-Object System.Drawing.Point($Tab1_row3_distance,85)
$Tab2_pkgmgrbutton5.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton5.Text = "Remove MS Store Apps"
$Tab2_pkgmgrbutton5.Add_Click({start-process powershell -verb runas {
Get-AppPackage | Remove-AppPackage
}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton5)

$Tab2_pkgmgrbutton6 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton6.Location = New-Object System.Drawing.Point($Tab1_row3_distance,120)
$Tab2_pkgmgrbutton6.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton6.Text = "Calculator"
$Tab2_pkgmgrbutton6.Add_Click({start-process powershell -verb runas {
    Get-AppxPackage -allusers *windowscalculator* | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register “$($_.InstallLocation)\AppXManifest.xml”}}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton6)

$Tab2_pkgmgrbutton7 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton7.Location = New-Object System.Drawing.Point($Tab1_row3_distance,155)
$Tab2_pkgmgrbutton7.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton7.Text = "win term"
$Tab2_pkgmgrbutton7.Add_Click({start-process powershell -verb runas {
Get-AppPackage | Remove-AppPackage
}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton7)

$Tab2_pkgmgrbutton8 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton8.Location = New-Object System.Drawing.Point($Tab1_row3_distance,190)
$Tab2_pkgmgrbutton8.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton8.Text = "wireless display adapter"
$Tab2_pkgmgrbutton8.Add_Click({start-process powershell -verb runas {
Get-AppPackage | Remove-AppPackage
}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton8)

$Tab2_pkgmgrbutton8 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton8.Location = New-Object System.Drawing.Point($Tab1_row3_distance,225)
$Tab2_pkgmgrbutton8.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton8.Text = "MS to-do"
$Tab2_pkgmgrbutton8.Add_Click({start-process powershell -verb runas {
Get-AppPackage | Remove-AppPackage
}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton8)




# TAB 3
#----------------------------------------------


$Tab2 = New-object System.Windows.Forms.Tabpage
$Tab2.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab2.Name = "Tab2" 
$Tab2.Text = "Laptop Migration" 
$FormTabControl.Controls.Add($Tab2)

# label
$Tab2_label1 = New-Object System.Windows.Forms.Label
$Tab2_label1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,10)
$Tab2_label1.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label1.AutoSize = $true
$Tab2_label1.ForeColor = "#000000"
$Tab2_label1.Text = ("Install")
$Tab2.Controls.Add($Tab2_label1)


#Add:
# Quick access backup
# Chrome, edge and firefox data backup and manual
# Start onedrive
# Find out way to enable backup locations with button
# Print out app list in log file





# TAB 4
#----------------------------------------------


$Tab2 = New-object System.Windows.Forms.Tabpage
$Tab2.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab2.Name = "Tab2" 
$Tab2.Text = "Laptop Setup" 
$FormTabControl.Controls.Add($Tab2)

# label
$Tab2_label1 = New-Object System.Windows.Forms.Label
$Tab2_label1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,10)
$Tab2_label1.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label1.AutoSize = $true
$Tab2_label1.ForeColor = "#000000"
$Tab2_label1.Text = ("Install")
$Tab2.Controls.Add($Tab2_label1)

#Add:
# VLC
# Image Glass
# Digidoc
# Real VNC viewer
# Powertoys
# Jabra
# Everything search
# Itunes



# TAB 5
#----------------------------------------------


$Tab2 = New-object System.Windows.Forms.Tabpage
$Tab2.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab2.Name = "Tab2" 
$Tab2.Text = "Desktop Setup" 
$FormTabControl.Controls.Add($Tab2)

# label
$Tab2_label1 = New-Object System.Windows.Forms.Label
$Tab2_label1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,10)
$Tab2_label1.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label1.AutoSize = $true
$Tab2_label1.ForeColor = "#000000"
$Tab2_label1.Text = ("Install")
$Tab2.Controls.Add($Tab2_label1)

# VLC







# TAB 10
#----------------------------------------------

$tab10 = New-object System.Windows.Forms.Tabpage
$Tab10.DataBindings.DefaultDataSourceUpdateMode = 0 
$tab10.Name = "tab10" 
$tab10.Text = "About” 
$FormTabControl.Controls.Add($tab10)

# window label
$tab10_label1 = New-Object System.Windows.Forms.Label
$tab10_label1.Location = New-Object System.Drawing.Point(10,10)
$tab10_label1.AutoSize = $true
$tab10_label1.Font = New-Object System.Drawing.Font('verdana',26,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$tab10_label1.ForeColor = "#000000"
$tab10_label1.Text =("JW OTB")
$tab10.Controls.Add($tab10_label1)


# window label
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,90)
$label2.AutoSize = $true
$label2.Font = New-Object System.Drawing.Font('verdana',12)
$label2.ForeColor = "#000000"
#$label2.backColor = "#b3b3b3"
$label2.Text = ("JW Open ToolBox is made by Efe Marko Güldere")
$tab10.Controls.Add($label2)

# link1
$link1 = New-Object System.Windows.Forms.LinkLabel
$link1.Location = New-Object System.Drawing.Point(10,150)
#$label.AutoSize = $true
#$label2.AutoSize.DisabledLinkColor = 'Blue'
$link1.VisitedLinkColor = 'Red'
$link1.LinkBehavior = 'HoverUnderline'
$link1.LinkColor = 'Navy'
$link1.Font = New-Object System.Drawing.Font('arial',10)
$link1.ForeColor = "#000000"
$link1.Text = ("Github Page")
$link1.add_click({explorer "https://github.com/Gren-95/JW-OTB"})
$tab10.Controls.Add($link1)

# link2
$link2 = New-Object System.Windows.Forms.LinkLabel
$link2.Location = New-Object System.Drawing.Point(110,150)
#$link2.AutoSize = $true
#$link2.AutoSize.DisabledLinkColor = 'Blue'
$link2.VisitedLinkColor = 'Red'
$link2.LinkBehavior = 'HoverUnderline'
$link2.LinkColor = 'Navy'
$link2.Font = New-Object System.Drawing.Font('arial',10)
$link2.ForeColor = "#000000"
#$link2.backColor = "#b3b3b3"
$link2.Text = ("Report issues")
$link2.add_click({explorer "https://github.com/Gren-95/JW-OTB/issues"})
$tab10.Controls.Add($link2)

# window label
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(590,400)
$label3.AutoSize = $true
$label3.Font = New-Object System.Drawing.Font('arial',8)
$label3.ForeColor = "#000000"
$label3.Text = ("Made in 2023")
$tab10.Controls.Add($label3)

# Initlize the form
$ApplicationForm.Add_Shown({$ApplicationForm.Activate()})
[void] $ApplicationForm.ShowDialog()