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
$ApplicationForm.Topmost = $true

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
$Tab1_appbutton1.Text = "Device Manager"
$Tab1_appbutton1.Add_Click({Start-Process hdwwiz.cpl})
$tab1.Controls.Add($Tab1_appbutton1)

$Tab1_appbutton2 = New-Object System.Windows.Forms.Button
$Tab1_appbutton2.Location = New-Object System.Drawing.Point($Tab1_row1_distance,85)
$Tab1_appbutton2.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton2.Text = "Display Properties"
$Tab1_appbutton2.Add_Click({Start-Process desk.cpl})
$tab1.Controls.Add($Tab1_appbutton2)

$Tab1_appbutton3 = New-Object System.Windows.Forms.Button
$Tab1_appbutton3.Location = New-Object System.Drawing.Point($Tab1_row1_distance,120)
$Tab1_appbutton3.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton3.Text = "Control Panel"
$Tab1_appbutton3.Add_Click({Start-Process control})
$tab1.Controls.Add($Tab1_appbutton3)

$Tab1_appbutton4 = New-Object System.Windows.Forms.Button
$Tab1_appbutton4.Location = New-Object System.Drawing.Point($Tab1_row1_distance,155)
$Tab1_appbutton4.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton4.Text = "Sound and Audio"
$Tab1_appbutton4.Add_Click({Start-Process mmsys.cpl})
$tab1.Controls.Add($Tab1_appbutton4)

$Tab1_appbutton5 = New-Object System.Windows.Forms.Button
$Tab1_appbutton5.Location = New-Object System.Drawing.Point($Tab1_row1_distance,190)
$Tab1_appbutton5.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton5.Text = "Network Properties"
$Tab1_appbutton5.Add_Click({Start-Process ncpa.cpl})
$tab1.Controls.Add($Tab1_appbutton5)

$Tab1_appbutton6 = New-Object System.Windows.Forms.Button
$Tab1_appbutton6.Location = New-Object System.Drawing.Point($Tab1_row1_distance,225)
$Tab1_appbutton6.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton6.Text = "Settings"
$Tab1_appbutton6.Add_Click({Start-Process ms-settings:})
$tab1.Controls.Add($Tab1_appbutton6)


$Tab1_appbutton9 = New-Object System.Windows.Forms.Button
$Tab1_appbutton9.Location = New-Object System.Drawing.Point($Tab1_row1_distance,260)
$Tab1_appbutton9.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton9.Text = "Calibrate touchscreen (Admin)"
$Tab1_appbutton9.Add_Click({Start-Process tabcal.exe})
$tab1.Controls.Add($Tab1_appbutton9)


$Tab1_appbutton10 = New-Object System.Windows.Forms.Button
$Tab1_appbutton10.Location = New-Object System.Drawing.Point($Tab1_row1_distance,295)
$Tab1_appbutton10.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton10.Text = "Add/Remove Applications"
$Tab1_appbutton10.Add_Click({Start-Process appwiz.cpl})
$tab1.Controls.Add($Tab1_appbutton10)


$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row1_distance,330)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Display Properties"
$Tab1_appbutton11.Add_Click({Start-Process desk.cpl})
$tab1.Controls.Add($Tab1_appbutton11)







$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row2_distance,50)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Update Windows"
$Tab1_appbutton11.Add_Click({
	Write-Host "Updating Windows"
	Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
	Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser
	Import-Module PSWindowsUpdate -Scope CurrentUser
	Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -Scope CurrentUser
})
$tab1.Controls.Add($Tab1_appbutton11)

$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row2_distance,85)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Update Office"
$Tab1_appbutton11.Add_Click({
	"$env:CommonProgramFiles\microsoft shared\ClickToRun\OfficeC2RClient.exe"
})
$tab1.Controls.Add($Tab1_appbutton11)

$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row2_distance,120)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Dell Command | Update"
$Tab1_appbutton11.Add_Click({Start-Process 
	'C:\Program Files\Dell\CommandUpdate\dcu-cli.exe'
})
$tab1.Controls.Add($Tab1_appbutton11)




$Tab1_appbutton7 = New-Object System.Windows.Forms.Button
$Tab1_appbutton7.Location = New-Object System.Drawing.Point($Tab1_row2_distance,225)
$Tab1_appbutton7.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton7.Text = "Printers"
$Tab1_appbutton7.Add_Click({Start-Process control printers})
$tab1.Controls.Add($Tab1_appbutton7)


$Tab1_appbutton8 = New-Object System.Windows.Forms.Button
$Tab1_appbutton8.Location = New-Object System.Drawing.Point($Tab1_row2_distance,260)
$Tab1_appbutton8.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton8.Text = "EERAKSRV02"
$Tab1_appbutton8.Add_Click({Start-Process \\eeraksrv02})
$tab1.Controls.Add($Tab1_appbutton8)




$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row3_distance,50)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Add to Domain"
$Tab1_appbutton11.Add_Click({
	Add-Computer -DomainName jwinc.jeld-wen.com
})
$tab1.Controls.Add($Tab1_appbutton11)

$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row3_distance,85)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Domain Panel"
$Tab1_appbutton11.Add_Click({Start-Process sysdm.cpl})
$tab1.Controls.Add($Tab1_appbutton11)

$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row3_distance,120)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Display Properties"
$Tab1_appbutton11.Add_Click({Start-Process desk.cpl})
$tab1.Controls.Add($Tab1_appbutton11)



$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row3_distance,225)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "JWWLAN"
$Tab1_appbutton11.Add_Click({
	netsh wlan connect ssid="JWWLAN" key="Jeld-Wen1!"
})
$tab1.Controls.Add($Tab1_appbutton11)

$Tab1_appbutton11 = New-Object System.Windows.Forms.Button
$Tab1_appbutton11.Location = New-Object System.Drawing.Point($Tab1_row3_distance,260)
$Tab1_appbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab1_appbutton11.Text = "Guest"
$Tab1_appbutton11.Add_Click({
	netsh wlan connect ssid="JWWLAN" key="Jeld-Wen1!"

})
$tab1.Controls.Add($Tab1_appbutton11)


























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
$Tab2_label2 = New-Object System.Windows.Forms.Label
$Tab2_label2.Location = New-Object System.Drawing.Point($Tab1_row2_distance,10)
$Tab2_label2.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label2.AutoSize = $true
$Tab2_label2.ForeColor = "#000000"
$Tab2_label2.Text = ("Remove")
$Tab2.Controls.Add($Tab2_label2)

# label
$Tab2_label3 = New-Object System.Windows.Forms.Label
$Tab2_label3.Location = New-Object System.Drawing.Point($Tab1_row3_distance,10)
$Tab2_label3.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label3.AutoSize = $true
$Tab2_label3.ForeColor = "#000000"
$Tab2_label3.Text = ("Pro Apps")
$Tab2.Controls.Add($Tab2_label3)

# label
$Tab2_label3 = New-Object System.Windows.Forms.Label
$Tab2_label3.Location = New-Object System.Drawing.Point($Tab1_row3_distance,190)
$Tab2_label3.Font = New-Object System.Drawing.Font('verdana',16)
$Tab2_label3.AutoSize = $true
$Tab2_label3.ForeColor = "#000000"
$Tab2_label3.Text = ("AppX")
$Tab2.Controls.Add($Tab2_label3)












$Tab2_installbutton1 = New-Object System.Windows.Forms.Button
$Tab2_installbutton1.Location = New-Object System.Drawing.Point($Tab1_row1_distance,50)
$Tab2_installbutton1.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton1.Text = "Digidoc4"
$Tab2_installbutton1.Add_Click({
	Invoke-WebRequest -Uri https://github.com/open-eid/DigiDoc4-Client/releases/download/v4.3.0/Digidoc4_Client-4.3.0.4438.x64.et-EE.qt.msi -OutFile "C:\Users\Public\Documents\digidoc.msi"
	Write-Host "downloaded"
	C:\Users\Public\Documents\digidoc.msi /quiet
	Write-Host "installed"

})
$tab2.Controls.Add($Tab2_installbutton1)

$Tab2_installbutton2 = New-Object System.Windows.Forms.Button
$Tab2_installbutton2.Location = New-Object System.Drawing.Point($Tab1_row1_distance,85)
$Tab2_installbutton2.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton2.Text = "Barcode Fonts"
$Tab2_installbutton2.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton2)

$Tab2_installbutton3 = New-Object System.Windows.Forms.Button
$Tab2_installbutton3.Location = New-Object System.Drawing.Point($Tab1_row1_distance,120)
$Tab2_installbutton3.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton3.Text = "Java"
$Tab2_installbutton3.Add_Click({Start-Process control})
$tab2.Controls.Add($Tab2_installbutton3)

$Tab2_installbutton4 = New-Object System.Windows.Forms.Button
$Tab2_installbutton4.Location = New-Object System.Drawing.Point($Tab1_row1_distance,155)
$Tab2_installbutton4.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton4.Text = "Cisco Anyconnect"
$Tab2_installbutton4.Add_Click({Start-Process mmsys.cpl})
$tab2.Controls.Add($Tab2_installbutton4)

$Tab2_installbutton5 = New-Object System.Windows.Forms.Button
$Tab2_installbutton5.Location = New-Object System.Drawing.Point($Tab1_row1_distance,190)
$Tab2_installbutton5.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton5.Text = "Jabra Connect"
$Tab2_installbutton5.Add_Click({Start-Process ncpa.cpl})
$tab2.Controls.Add($Tab2_installbutton5)

$Tab2_installbutton6 = New-Object System.Windows.Forms.Button
$Tab2_installbutton6.Location = New-Object System.Drawing.Point($Tab1_row1_distance,225)
$Tab2_installbutton6.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton6.Text = "Adobe Acrobat"
$Tab2_installbutton6.Add_Click({Start-Process ms-settings:})
$tab2.Controls.Add($Tab2_installbutton6)


$Tab2_installbutton7 = New-Object System.Windows.Forms.Button
$Tab2_installbutton7.Location = New-Object System.Drawing.Point($Tab1_row1_distance,260)
$Tab2_installbutton7.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton7.Text = "Add Active Directory Users and Computers"
$Tab2_installbutton7.Add_Click({Start-Process tabcal.exe})
$tab2.Controls.Add($Tab2_installbutton7)


$Tab2_installbutton8 = New-Object System.Windows.Forms.Button
$Tab2_installbutton8.Location = New-Object System.Drawing.Point($Tab1_row1_distance,295)
$Tab2_installbutton8.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton8.Text = "MS Office"
$Tab2_installbutton8.Add_Click({Start-Process appwiz.cpl})
$tab2.Controls.Add($Tab2_installbutton8)


$Tab2_installbutton9 = New-Object System.Windows.Forms.Button
$Tab2_installbutton9.Location = New-Object System.Drawing.Point($Tab1_row1_distance,330)
$Tab2_installbutton9.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton9.Text = "Firefox"
$Tab2_installbutton9.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton9)


$Tab2_installbutton10 = New-Object System.Windows.Forms.Button
$Tab2_installbutton10.Location = New-Object System.Drawing.Point($Tab1_row1_distance,365)
$Tab2_installbutton10.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton10.Text = "Powertoys"
$Tab2_installbutton10.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton10)


$Tab2_installbutton11 = New-Object System.Windows.Forms.Button
$Tab2_installbutton11.Location = New-Object System.Drawing.Point($Tab1_row1_distance,400)
$Tab2_installbutton11.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton11.Text = "WSL"
$Tab2_installbutton11.Add_Click({start-process powershell -verb runas {
	Enable-WindowsOptionalFeature -Online -FeatureName "VirtualMachinePlatform" -All
	Enable-WindowsOptionalFeature -Online -FeatureName "Microsoft-Windows-Subsystem-Linux" -All
	Write-Host "WSL is now installed and configured. Please Reboot before using."
	}})
$tab2.Controls.Add($Tab2_installbutton11)













$Tab2_installbutton1 = New-Object System.Windows.Forms.Button
$Tab2_installbutton1.Location = New-Object System.Drawing.Point($Tab1_row2_distance,50)
$Tab2_installbutton1.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton1.Text = "Java"
$Tab2_installbutton1.Add_Click({Start-Process hdwwiz.cpl})
$tab2.Controls.Add($Tab2_installbutton1)

$Tab2_installbutton2 = New-Object System.Windows.Forms.Button
$Tab2_installbutton2.Location = New-Object System.Drawing.Point($Tab1_row2_distance,85)
$Tab2_installbutton2.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton2.Text = "Adobe Acrobat"
$Tab2_installbutton2.Add_Click({Start-Process desk.cpl})
$tab2.Controls.Add($Tab2_installbutton2)

$Tab2_installbutton3 = New-Object System.Windows.Forms.Button
$Tab2_installbutton3.Location = New-Object System.Drawing.Point($Tab1_row2_distance,120)
$Tab2_installbutton3.Size = New-Object System.Drawing.Size(200,35)
$Tab2_installbutton3.Text = "MS Office"
$Tab2_installbutton3.Add_Click({Start-Process control})
$tab2.Controls.Add($Tab2_installbutton3)













$Tab2_pkgmgrbutton4 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton4.Location = New-Object System.Drawing.Point($Tab1_row3_distance,225)
$Tab2_pkgmgrbutton4.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton4.Text = "Reinstall MS Store"
$Tab2_pkgmgrbutton4.Add_Click({start-process powershell -verb runas {
Get-AppxPackage -allusers Microsoft.WindowsStore | ForEach-Object {Add-AppxPackage -DisableDevelopmentMode -Register “$($_.InstallLocation)\AppXManifest.xml”}
}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton4)

$Tab2_pkgmgrbutton5 = New-Object System.Windows.Forms.Button
$Tab2_pkgmgrbutton5.Location = New-Object System.Drawing.Point($Tab1_row3_distance,260)
$Tab2_pkgmgrbutton5.Size = New-Object System.Drawing.Size (200,35)
$Tab2_pkgmgrbutton5.Text = "Remove MS Store Apps"
$Tab2_pkgmgrbutton5.Add_Click({start-process powershell -verb runas {
Get-AppPackage | Remove-AppPackage
}})
$Tab2.Controls.Add($Tab2_pkgmgrbutton5)


$Tab2_probutton1 = New-Object System.Windows.Forms.Button
$Tab2_probutton1.Location = New-Object System.Drawing.Point($Tab1_row3_distance,50)
$Tab2_probutton1.Size = New-Object System.Drawing.Size (200,35)
$Tab2_probutton1.Text = "Install User Manager"
$Tab2_probutton1.Add_Click({start-process powershell -verb runas {
mkdir 'C:\Program Files\usermanager'
Start-BitsTransfer -Source "https://github.com/proviq/AccountManagement/raw/master/lusrmgr/bin/Release/lusrmgr.exe" -Destination "C:\Program Files\usermanager"
$SourceFilePath = "C:\Program Files\usermanager\lusrmgr.exe"
$ShortcutPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\lusrmgr.lnk"
$WScriptObj = New-Object -ComObject ("WScript.Shell")
$shortcut = $WscriptObj.CreateShortcut($ShortcutPath)
$shortcut.TargetPath = $SourceFilePath
$shortcut.Save()
}})
$Tab2.Controls.Add($Tab2_probutton1)

$Tab2_probutton2 = New-Object System.Windows.Forms.Button
$Tab2_probutton2.Location = New-Object System.Drawing.Point($Tab1_row3_distance,85)
$Tab2_probutton2.Size = New-Object System.Drawing.Size (200,35)
$Tab2_probutton2.Text = "User Manager"
$Tab2_probutton2.Add_Click({& 'C:\Program Files\usermanager\lusrmgr.exe'})
$Tab2.Controls.Add($Tab2_probutton2)










# TAB 3
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



# TAB 4
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