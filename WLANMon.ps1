##############################################################################################
#
# Windows Netsh WLANMon Powershell script BETA v0.02
#
# The inspiration and starting place for this script was Nigel Bowden's powershell script
# http://wifinigel.blogspot.com/2016/09/getting-data-out-of-windows-netsh-wlan.html
# which was Inspired by Matt Frederick's blog post: 
# https://finesine.com/2016/09/17/using-netsh-wlan-show-interfaces-to-monitor-associationroaming/
#
# Note: to run Powershell scripts, you will most likely need to update
#       the execution policy on your machine. Open a Powershell window
#       as an adminsitrator on your machine and temporarily set the policy
#       to unrestricted with this command:
#
#		Set-ExecutionPolicy Unrestricted -scope Process
#
#       Once your powershell window is closed, this policy change is no
#       longer in effect and your machine will return to the previous
#       policy. You can check your execution policy by running the 
#       following command in Powershell:
#
#       Get-ExecutionPolicy
#
# Versions:
# BETA V0.01 - Oringal
# BETA V0.02 - Chnaged log output to ASCII, SO Log csv file will open correclty in Excel
##########################################################################################


#Define and buid window GUI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$WLANMon                        = New-Object system.Windows.Forms.Form
$WLANMon.ClientSize             = '655,804'
$WLANMon.BackColor              = "#000000"
$WLANMon.text                   = "Netsh WLANMon - Beta 0.02"
$WLANMon.TopMost                = $false
$WLANMon.location        		= New-Object System.Drawing.Point(10,10)

$Requirements                    = New-Object system.Windows.Forms.Groupbox
$Requirements.height             = 73
$Requirements.width              = 160
$Requirements.text               = "Requirements"
$Requirements.ForeColor          = "#ffffff"
$Requirements.location           = New-Object System.Drawing.Point(10,10)

$LogSettings                     = New-Object system.Windows.Forms.Groupbox
$LogSettings.height             = 73
$LogSettings.width              = 330
$LogSettings.text               = "Log Settings"
$LogSettings.ForeColor          = "#ffffff"
$LogSettings.location           = New-Object System.Drawing.Point(190,10)

$WinAdaptorGroup                 = New-Object system.Windows.Forms.Groupbox
$WinAdaptorGroup.height          = 76
$WinAdaptorGroup.width           = 635
$WinAdaptorGroup.text            = "Wi-Fi Adaptor"
$WinAdaptorGroup.ForeColor       = "#ffffff"
$WinAdaptorGroup.location        = New-Object System.Drawing.Point(10,93)

$NetworkGroup                    = New-Object system.Windows.Forms.Groupbox
$NetworkGroup.height             = 106
$NetworkGroup.width              = 635
$NetworkGroup.text               = "Network"
$NetworkGroup.ForeColor           = "#ffffff"
$NetworkGroup.location           = New-Object System.Drawing.Point(10,179)

$RSSIGroup                       = New-Object system.Windows.Forms.Groupbox
$RSSIGroup.height                = 106
$RSSIGroup.width                 = 635
$RSSIGroup.text                  = "RSSI"
$RSSIGroup.ForeColor           = "#ffffff"
$RSSIGroup.location              = New-Object System.Drawing.Point(10,295)

$RoamGroup                       = New-Object system.Windows.Forms.Groupbox
$RoamGroup.height                = 383
$RoamGroup.width                 = 635
$RoamGroup.text                  = "Roaming Log"
$RoamGroup.ForeColor           = "#ffffff"
$RoamGroup.location              = New-Object System.Drawing.Point(10,411)

$AdaptorName                     = New-Object system.Windows.Forms.Label
$AdaptorName.text                = "AdaptorName"
$AdaptorName.AutoSize            = $true
$AdaptorName.width               = 25
$AdaptorName.height              = 10
$AdaptorName.location            = New-Object System.Drawing.Point(12,18)
$AdaptorName.Font                = 'Microsoft Sans Serif,10'
$AdaptorName.ForeColor           = "#ffffff"

$MACLabel                        = New-Object system.Windows.Forms.Label
$MACLabel.text                   = "MAC:"
$MACLabel.AutoSize               = $true
$MACLabel.width                  = 25
$MACLabel.height                 = 10
$MACLabel.location               = New-Object System.Drawing.Point(12,45)
$MACLabel.Font                   = 'Microsoft Sans Serif,10'
$MACLabel.ForeColor              = "#ffffff"

$MAC                             = New-Object system.Windows.Forms.Label
$MAC.text                        = ""
$MAC.AutoSize                    = $true
$MAC.width                       = 25
$MAC.height                      = 10
$MAC.location                    = New-Object System.Drawing.Point(55,45)
$MAC.Font                        = 'Microsoft Sans Serif,10'
$MAC.ForeColor                   = "#ffffff"

$StartButton                     = New-Object system.Windows.Forms.Button
$StartButton.BackColor           = "#b8e986"
$StartButton.text                = "Start"
$StartButton.width               = 100
$StartButton.height              = 30
$StartButton.location            = New-Object System.Drawing.Point(535,20)
$StartButton.Font                = 'Microsoft Sans Serif,10'

$Connected                       = New-Object system.Windows.Forms.Panel
$Connected.height                = 16
$Connected.width                 = 16
$Connected.BackColor             = "#d0021b"
$Connected.location              = New-Object System.Drawing.Point(614,58)

$ConnectedLabel                  = New-Object system.Windows.Forms.Label
$ConnectedLabel.text             = "Connected:"
$ConnectedLabel.AutoSize         = $true
$ConnectedLabel.width            = 25
$ConnectedLabel.height           = 10
$ConnectedLabel.location         = New-Object System.Drawing.Point(540,58)
$ConnectedLabel.Font             = 'Microsoft Sans Serif,10'
$ConnectedLabel.ForeColor        = "#ffffff"

$RadioLabel                      = New-Object system.Windows.Forms.Label
$RadioLabel.text                 = "Radio Type:"
$RadioLabel.AutoSize             = $true
$RadioLabel.width                = 25
$RadioLabel.height               = 10
$RadioLabel.location             = New-Object System.Drawing.Point(272,45)
$RadioLabel.Font                 = 'Microsoft Sans Serif,10'
$RadioLabel.ForeColor            = "#ffffff"

$RadioType                       = New-Object system.Windows.Forms.Label
$RadioType.text                  = ""
$RadioType.AutoSize              = $true
$RadioType.width                 = 25
$RadioType.height                = 10
$RadioType.location              = New-Object System.Drawing.Point(348,45)
$RadioType.Font                  = 'Microsoft Sans Serif,10'
$RadioType.ForeColor             = "#ffffff"

$RSSIDBLabel                     = New-Object system.Windows.Forms.Label
$RSSIDBLabel.text                = "Signal(dBm)"
$RSSIDBLabel.AutoSize            = $true
$RSSIDBLabel.width               = 25
$RSSIDBLabel.height              = 10
$RSSIDBLabel.location            = New-Object System.Drawing.Point(37,19)
$RSSIDBLabel.Font                = 'Microsoft Sans Serif,10'
$RSSIDBLabel.ForeColor           = "#ffffff"

$RSSIPercentLabel                = New-Object system.Windows.Forms.Label
$RSSIPercentLabel.text           = "Signal (%)"
$RSSIPercentLabel.AutoSize       = $true
$RSSIPercentLabel.width          = 25
$RSSIPercentLabel.height         = 10
$RSSIPercentLabel.location       = New-Object System.Drawing.Point(201,18)
$RSSIPercentLabel.Font           = 'Microsoft Sans Serif,12'
$RSSIPercentLabel.ForeColor      = "#ffffff"

$SSIDLabel                       = New-Object system.Windows.Forms.Label
$SSIDLabel.text                  = "SSID:"
$SSIDLabel.AutoSize              = $true
$SSIDLabel.width                 = 25
$SSIDLabel.height                = 10
$SSIDLabel.location              = New-Object System.Drawing.Point(15,18)
$SSIDLabel.Font                  = 'Microsoft Sans Serif,12'
$SSIDLabel.ForeColor             = "#ffffff"

$BSSIDLabel                      = New-Object system.Windows.Forms.Label
$BSSIDLabel.text                 = "BSSID:"
$BSSIDLabel.AutoSize             = $true
$BSSIDLabel.width                = 25
$BSSIDLabel.height               = 10
$BSSIDLabel.location             = New-Object System.Drawing.Point(12,63)
$BSSIDLabel.Font                 = 'Microsoft Sans Serif,12'
$BSSIDLabel.ForeColor            = "#ffffff"

$SSID                            = New-Object system.Windows.Forms.Label
$SSID.text                       = ""
$SSID.AutoSize                   = $true
$SSID.width                      = 25
$SSID.height                     = 10
$SSID.location                   = New-Object System.Drawing.Point(77,18)
$SSID.Font                       = 'Microsoft Sans Serif,12'
$SSID.ForeColor                  = "#7ed321"

$BSSID                           = New-Object system.Windows.Forms.Label
$BSSID.text                      = ""
$BSSID.AutoSize                  = $true
$BSSID.width                     = 25
$BSSID.height                    = 10
$BSSID.location                  = New-Object System.Drawing.Point(73,61)
$BSSID.Font                      = 'Microsoft Sans Serif,16'
$BSSID.ForeColor                 = "#f5a623"

$SignaldB                        = New-Object system.Windows.Forms.Label
$SignaldB.text                   = ""
$SignaldB.AutoSize               = $true
$SignaldB.width                  = 25
$SignaldB.height                 = 10
$SignaldB.location               = New-Object System.Drawing.Point(43,49)
$SignaldB.Font                   = 'Microsoft Sans Serif,26'
$SignaldB.ForeColor              = "#7ed321"

$SignalPercent                   = New-Object system.Windows.Forms.Label
$SignalPercent.text              = ""
$SignalPercent.AutoSize          = $true
$SignalPercent.width             = 25
$SignalPercent.height            = 10
$SignalPercent.location          = New-Object System.Drawing.Point(198,48)
$SignalPercent.Font              = 'Microsoft Sans Serif,26'
$SignalPercent.ForeColor         = "#7ed321"

$RoamingList                     = New-Object system.Windows.Forms.ListView
$RoamingList.View                  	 = 'Details'
$RoamingList.width               = 579
$RoamingList.height              = 307
$RoamingList.Columns.Add('#')
$RoamingList.Columns.Add('Time')
$RoamingList.Columns.Add('From')
$RoamingList.Columns.Add('To')
$RoamingList.location            = New-Object System.Drawing.Point(22,55)


$AuthentictionLabel              = New-Object system.Windows.Forms.Label
$AuthentictionLabel.text         = "Authentication:"
$AuthentictionLabel.AutoSize     = $true
$AuthentictionLabel.width        = 25
$AuthentictionLabel.height       = 10
$AuthentictionLabel.location     = New-Object System.Drawing.Point(244,18)
$AuthentictionLabel.Font         = 'Microsoft Sans Serif,12'
$AuthentictionLabel.ForeColor    = "#ffffff"

$Authentiction                   = New-Object system.Windows.Forms.Label
$Authentiction.text              = ""
$Authentiction.AutoSize          = $true
$Authentiction.width             = 25
$Authentiction.height            = 10
$Authentiction.location          = New-Object System.Drawing.Point(348,18)
$Authentiction.Font              = 'Microsoft Sans Serif,12'
$Authentiction.ForeColor         = "#7ed321"

$CipherLabel                     = New-Object system.Windows.Forms.Label
$CipherLabel.text                = "Cipher:"
$CipherLabel.AutoSize            = $true
$CipherLabel.width               = 25
$CipherLabel.height              = 10
$CipherLabel.location            = New-Object System.Drawing.Point(485,18)
$CipherLabel.Font                = 'Microsoft Sans Serif,12'
$CipherLabel.ForeColor           = "#ffffff"

$Cipher                          = New-Object system.Windows.Forms.Label
$Cipher.text                     = ""
$Cipher.AutoSize                 = $true
$Cipher.width                    = 25
$Cipher.height                   = 10
$Cipher.location                 = New-Object System.Drawing.Point(547,18)
$Cipher.Font                     = 'Microsoft Sans Serif,12'
$Cipher.ForeColor                = "#7ed321"

$ChanneLabel                     = New-Object system.Windows.Forms.Label
$ChanneLabel.text                = "Channel:"
$ChanneLabel.AutoSize            = $true
$ChanneLabel.width               = 25
$ChanneLabel.height              = 10
$ChanneLabel.location            = New-Object System.Drawing.Point(480,61)
$ChanneLabel.Font                = 'Microsoft Sans Serif,12'
$ChanneLabel.ForeColor           = "#ffffff"

$Channel                         = New-Object system.Windows.Forms.Label
$Channel.text                    = ""
$Channel.AutoSize                = $true
$Channel.width                   = 25
$Channel.height                  = 10
$Channel.location                = New-Object System.Drawing.Point(550,50)
$Channel.Font                    = 'Microsoft Sans Serif,24'
$Channel.ForeColor               = "#ffffff"

$TXLable                         = New-Object system.Windows.Forms.Label
$TXLable.text                    = "TX Data Rate:"
$TXLable.AutoSize                = $true
$TXLable.width                   = 25
$TXLable.height                  = 10
$TXLable.location                = New-Object System.Drawing.Point(334,19)
$TXLable.Font                    = 'Microsoft Sans Serif,12'
$TXLable.ForeColor               = "#ffffff"

$RXLabel                         = New-Object system.Windows.Forms.Label
$RXLabel.text                    = "RX Date Rate:"
$RXLabel.AutoSize                = $true
$RXLabel.width                   = 25
$RXLabel.height                  = 10
$RXLabel.location                = New-Object System.Drawing.Point(487,20)
$RXLabel.Font                    = 'Microsoft Sans Serif,12'
$RXLabel.ForeColor               = "#ffffff"

$TXDataRate                      = New-Object system.Windows.Forms.Label
$TXDataRate.text                 = ""
$TXDataRate.AutoSize             = $true
$TXDataRate.width                = 25
$TXDataRate.height               = 10
$TXDataRate.location             = New-Object System.Drawing.Point(351,49)
$TXDataRate.Font                 = 'Microsoft Sans Serif,26'
$TXDataRate.ForeColor            = "#7ed321"

$RXDataRate                      = New-Object system.Windows.Forms.Label
$RXDataRate.text                 = ""
$RXDataRate.AutoSize             = $true
$RXDataRate.width                = 25
$RXDataRate.height               = 10
$RXDataRate.location             = New-Object System.Drawing.Point(504,49)
$RXDataRate.Font                 = 'Microsoft Sans Serif,26'
$RXDataRate.ForeColor            = "#7ed321"

$SignalReqlabel                  = New-Object system.Windows.Forms.Label
$SignalReqlabel.text             = "Signal:"
$SignalReqlabel.AutoSize         = $true
$SignalReqlabel.width            = 25
$SignalReqlabel.height           = 10
$SignalReqlabel.location         = New-Object System.Drawing.Point(15,18)
$SignalReqlabel.Font             = 'Microsoft Sans Serif,12'
$SignalReqlabel.ForeColor        = "#ffffff"

$DataRateReqLabel                = New-Object system.Windows.Forms.Label
$DataRateReqLabel.text           = "Data Rate:"
$DataRateReqLabel.AutoSize       = $true
$DataRateReqLabel.width          = 25
$DataRateReqLabel.height         = 10
$DataRateReqLabel.location       = New-Object System.Drawing.Point(15,43)
$DataRateReqLabel.Font           = 'Microsoft Sans Serif,12'
$DataRateReqLabel.ForeColor      = "#ffffff"

$SignalRequirment                = New-Object system.Windows.Forms.TextBox
$SignalRequirment.multiline      = $false
$SignalRequirment.text           = "-65"
$SignalRequirment.width          = 30
$SignalRequirment.height         = 17
$SignalRequirment.Anchor         = 'left'
$SignalRequirment.location       = New-Object System.Drawing.Point(100,18)
$SignalRequirment.Font           = 'Microsoft Sans Serif,8'

$DataRateReqirment               = New-Object system.Windows.Forms.TextBox
$DataRateReqirment.multiline     = $false
$DataRateReqirment.text          = "24"
$DataRateReqirment.width         = 30
$DataRateReqirment.height        = 20
$DataRateReqirment.location      = New-Object System.Drawing.Point(100,43)
$DataRateReqirment.Font          = 'Microsoft Sans Serif,8'

$RoamCountLabel                  = New-Object system.Windows.Forms.Label
$RoamCountLabel.text             = "RoamCount:"
$RoamCountLabel.AutoSize         = $true
$RoamCountLabel.width            = 25
$RoamCountLabel.height           = 10
$RoamCountLabel.location         = New-Object System.Drawing.Point(27,20)
$RoamCountLabel.Font             = 'Microsoft Sans Serif,12'
$RoamCountLabel.ForeColor        = "#ffffff"

$RoamCount                       = New-Object system.Windows.Forms.Label
$RoamCount.text                  = "0"
$RoamCount.AutoSize              = $true
$RoamCount.width                 = 25
$RoamCount.height                = 10
$RoamCount.location              = New-Object System.Drawing.Point(126,20)
$RoamCount.Font                  = 'Microsoft Sans Serif,14,style=Bold'
$RoamCount.ForeColor             = "#ffffff"

$LogCheckbox					 = New-Object System.Windows.Forms.Checkbox 
$LogCheckbox.Location 			 = New-Object System.Drawing.Size(15,15) 
$LogCheckbox.Text 				 = "Enable Logging"

$FileNameLabel               	 = New-Object system.Windows.Forms.Label
$FileNameLabel.text          	 = "File Name"
$FileNameLabel.AutoSize      	 = $true
$FileNameLabel.width        	 = 25
$FileNameLabel.height        	 = 10
$FileNameLabel.location      	 = New-Object System.Drawing.Point(15,43)
$FileNameLabel.Font          	 = 'Microsoft Sans Serif,12'
$FileNameLabel.ForeColor     	 = "#ffffff"

$LogFileName           			 = New-Object system.Windows.Forms.TextBox
$LogFileName.multiline   	 	 = $false
$LogFileName.text        		 = "WLANMonLog"
$LogFileName.width				 = 220
$LogFileName.height       		 = 20
$LogFileName.location    		 = New-Object System.Drawing.Point(100,43)
$LogFileName.Font        		 = 'Microsoft Sans Serif,8'


$WLANMon.controls.AddRange(@($TitleLable,$WinAdaptorGroup,$RSSIGroup,$NetworkGroup,$RoamGroup,$Requirements, $LogSettings,$Connected,$ConnectedLabel,$StartButton))
$WinAdaptorGroup.controls.AddRange(@($AdaptorName,$MACLabel,$MAC,$RadioLabel,$RadioType))
$RSSIGroup.controls.AddRange(@($RSSIDBLabel,$RSSIPercentLabel,$SignaldB,$SignalPercent,$TXLable,$RXLabel,$TXDataRate,$RXDataRate))
$NetworkGroup.controls.AddRange(@($SSIDLabel,$BSSIDLabel,$SSID,$BSSID,$AuthentictionLabel,$Authentiction,$CipherLabel,$Cipher,$ChanneLabel,$Channel))
$RoamGroup.controls.AddRange(@($RoamingList,$RoamCountLabel,$RoamCount))
$Requirements.controls.AddRange(@($SignalReqlabel,$DataRateReqLabel,$SignalRequirment,$DataRateReqirment))
$LogSettings.controls.AddRange(@($LogCheckbox,$FileNameLabel,$LogFileName))


#Init variables
   $LoggingEnabled = $false
   $script:CancelLoop = $false
   $Adaptor = ''
   $MACAdd = ''
   $SSIDText = ''
   $BSSIDText = ''
   $NetType = ''
   $RadType = '' 
   $Auth = ''
   $CipherText = ''
   $Chan = ''
   $Sig = ''
   $TxRate = ''
   $RxRate = ''   
   $dBmSig =''
   $SignalLevelPercent = ''
   $OldBSSID = ''

#Start button is click
$StartButton.Add_Click({  

# Define loop wait time in secs
$SleepInterval = 1

#Init variables
$RoamNum = 0
$CurrentTime = Get-Date
$name = $LogFileName.text
$day = ($CurrentTime -split "/")[0].Trim()
$month = ($CurrentTime -split "/")[1].Trim()
$year = ($CurrentTime -split "/")[2].substring(0,4)
$hour = ($CurrentTime -split ":")[0].substring(11)
$min = ($CurrentTime -split ":")[1].Trim()
$sec = ($CurrentTime -split ":")[-1]


#Start button control
If ($StartButton.text -eq "Start") {
   $timestamp = "$day-$month-$year-$hour.$min.$sec"
   $filename = "$name-$timestamp.csv"
   $StartButton.text = "End"
   $StartButton.BackColor = "#d0021b"
   $script:CancelLoop = $false
   $OldBSSID = ''
   If ($LogCheckbox.Checked -eq $true) {
      $LoggingEnabled = $true
      $headers = "CurrentTime, Name, Description, GUID, MAC, State, SSID, BSSID, NetworkType, RadioType, Authentication, Cipher, Connection, Channel, RecRate, TransRate, SignalLevelPercent, SignalLeveldBm, Profile"
      $headers | Out-File -FilePath $filename -Encoding ascii	  
   }
}
else {
   $StartButton.text = "Start"
   $StartButton.BackColor = "#b8e986"
   $script:CancelLoop = $true
}

#Execute the netsh commmand
$output = netsh.exe wlan show interfaces
 

# Start Loop
Do{

  #Run netsh command to get wirelss profile info
  $output = netsh.exe wlan show interfaces

  # Get time to time-stamp entry
  $CurrentTime = Get-Date

  # Name
  $Name_line = $output | Select-String -Pattern 'Name'
  $Name = ($Name_line -split ":")[-1].Trim()

  # Description
  $Description_line = $output | Select-String -Pattern 'Description'
  $Adaptor = ($Description_line -split ":")[-1].Trim()
  $AdaptorName.text = $Adaptor

  # GUID
  $GUID_line = $output | Select-String -Pattern 'GUID'
  $GUID = ($GUID_line -split ":")[-1].Trim()

  # Physical Address
  $Physical_line = $output | Select-String -Pattern 'Physical'
  $MACAdd = ($Physical_line -split ":", 2)[-1].Trim()
  $MAC.text = $MACAdd

  # State
  $State_line = $output | Select-String -Pattern 'State'
  $State = ($State_line -split ":")[-1].Trim()

  if ($State -eq 'connected') {
  
    $Connected.BackColor             = "#7ed321"

    # SSID
    $SSID_line = $output | Select-String 'SSID'| select -First 1
    $SSIDText = ($SSID_line -split ":")[-1].Trim()
    $SSID.text = $SSIDText

    # BSSID
    $BSSID_line = $output | Select-String -Pattern 'BSSID'
    $BSSIDText = ($BSSID_line -split ":", 2)[-1].Trim()
    $BSSID.text = $BSSIDText

    # NetworkType
    $NetworkType_line = $output | Select-String -Pattern 'Network type'
    $NetworkType = ($NetworkType_line -split ":")[-1].Trim()

    # RadioType
    $RadioType_line = $output | Select-String -Pattern 'Radio type'
    $RadType = ($RadioType_line -split ":")[-1].Trim()
    $RadioType.text = $RadType

    # Authentication
    $Authentication_line = $output | Select-String -Pattern 'Authentication'
    $Auth = ($Authentication_line -split ":")[-1].Trim()
    $Authentiction.text = $Auth

    # Cipher
    $Cipher_line = $output | Select-String -Pattern 'Cipher'
    $CipherText = ($Cipher_line -split ":")[-1].Trim()
    $Cipher.text = $CipherText

    # Connection mode
    $Connection_line = $output | Select-String -Pattern 'Connection mode'
    $Connection = ($Connection_line -split ":")[-1].Trim()

    # Channel
    $Channel_line = $output | Select-String -Pattern 'Channel'
    $Chan = ($Channel_line -split ":")[-1].Trim()
    $Channel.text = $Chan

    # Receive Rate
    $RecRate_line = $output | Select-String -Pattern 'Receive rate'
    $RxRate = ($RecRate_line -split ":")[-1].Trim()
    $RXDataRate.text = $RxRate
	
	
    # Transmit Rate
    $TransRate_line = $output | Select-String -Pattern 'Transmit rate'
    $TxRate = ($TransRate_line -split ":")[-1].Trim()
    $TXDataRate.text = $TxRate

    #Evaluate transmit and recieve data rate against requirements
	if([int]$RXDataRate.text -lt [int]$DataRateReqirment.text) {
		$RXDataRate.ForeColor = "#d0021b"
	} else {
		$RXDataRate.ForeColor = "#b8e986"
	}

	if([int]$TXDataRate.text -lt [int]$DataRateReqirment.text) {
		$TXDataRate.ForeColor = "#d0021b"
	} else {
		$TXDataRate.ForeColor = "#b8e986"
	}
	
    # Signal (%)
    $SignalLevelPercent_line = $output | Select-String -Pattern 'Signal'
    $SignalLevelPercent = ($SignalLevelPercent_line -split ":")[-1].Trim()	
    $SignalPercent.text = $SignalLevelPercent

	# Signal (dBm)
    $SignalLevelPercent_trimmed = $SignalLevelPercent.TrimEnd('%')
    $dBmSig = (([int]$SignalLevelPercent_trimmed)/2) - 100
	$SignaldB.text = $dBmSig

    #Evaluated signal strength against requirment 
	if($dBmSig -lt [int]$SignalRequirment.text) {
		$SignaldB.ForeColor = "#d0021b"
		$SignalPercent.ForeColor = "#d0021b"
	}
	else{
		$SignaldB.ForeColor = "#b8e986"
		$SignalPercent.ForeColor = "#b8e986"
	}
	
    # Signal (dBm)
    $SignalLevelPercent_trimmed = $SignalLevelPercent.TrimEnd('%')
    $dBmSig = (([int]$SignalLevelPercent_trimmed)/2) - 100
	$SignaldB.text = $dBmSig

    # Profile
    $Profile_line = $output | Select-String -Pattern 'Profile'
    $Profile = ($Profile_line -split ":")[-1].Trim()
	
  #Handle roaming	
  if (-NOT ($BSSID.text -eq $OldBSSID)) {
     if (-NOT ($OldBSSID -eq '')) {
        $RoamNum = $RoamNum + 1
	    $RoamCount.text = $RoamNum
	    $CurrentBSS = $BSSID.text
	    $item1 = New-Object System.Windows.Forms.ListViewItem($RoamNum)
	    $item1.SubItems.Add("$CurrentTime")
        $item1.SubItems.Add($OldBSSID)
	    $item1.SubItems.Add($CurrentBSS)
	    $RoamingList.Items.Add($item1)
	    $RoamingList.AutoResizeColumns(2)
	 }
  }  
}
else {
   $Connected.BackColor             = "#d0021b"
   $Authentiction.text = ''
   $Cipher.text = ''
   $Channel.text =''
   $TXDataRate.text =''
   $RXDataRate.text = ''
   $SignalPercent.text = ''
   $SignaldB.text = ''
}

#Write log file
If ($LogCheckbox.checked) {
   $logline = "$CurrentTime, $Name, $Adaptor, $GUID, $MACAdd, $State, $SSIDText, $BSSIDText, $NetworkType, $RadType, $Auth, $CipherText, $Connection, $Chan, $RxRate, $TxRate, $SignalLevelPercent, $dBmSig, $Profile"

   if ($LoggingEnabled -eq $true) {
      $logline | Out-File -append -FilePath $filename -Encoding ascii
   } else {
      $headers | Out-File -FilePath $filename -Encoding ascii
      $logline | Out-File -append -FilePath $filename -Encoding ascii

	  $LoggingEnabled = $true
   }
}
else {
   $LoggingEnabled = $false
}

#Roaming control
$OldBSSID = $BSSID.text

#give time to external events out side of loop 
[System.Windows.Forms.Application]::DoEvents()
  
#Exit loop control  
If($script:CancelLoop -eq $true) {
     break;
}

#loop sleep
Start-Sleep -s $SleepInterval

}
Until (0)

})

#Break loop is form close is clicked
$WLANMon.Add_FormClosing({ 
    $script:CancelLoop = $true
})


#display GUI
$WLANMon.ShowDialog()