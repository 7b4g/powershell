# Get info from PC: FIO, programs (ex. office 2013), processor, monitors, RAM, hard drive, PC name, PC model, OS
# Before stating, check all comments and values

# Params

$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$username = $env:UserName
$computerName = $env:computername

# For network folder save
# $corePath = "YOUR PATH TO NETWORK FOLDER"
# $savePath = $corePath + $computerName + '.txt'
 
# For local save
$userPath = [Environment]::GetFolderPath("Desktop")
$savePath = $userPath + "\" + $computerName + ".txt"

$skip = '///////////////////////////////////////////////'

# Checks

# Path available? (USE WITH NETWORK FOLDER)
# $miniPath = "YOUR PATH TO NETWORK FOLDER"
# $miniPathTest = [System.IO.Directory]::Exists($miniPath)
# if ($miniPathTest -eq "False"){return}

# File exist? If exist, doesnt create new
$fileCheck = Test-Path $savePath
if ($fileCheck -eq "True"){return}

# Core func
# Full name in domain
$fullName = ([adsi]"WinNT://$env:userdomain/$env:username,user").fullname

# Find Program Name, installed
$win32Product = gwmi win32_product

#$check = $win32Product | where-object name -match "PROGRAM NAME"
# If not installed one, then another installed
#if ($check -eq $null){$check = "PROGRAM NAME"} else {$check = "THEN INSTALLED ANOTHER PROGRAM"}

# Office as example
$office = $win32Product | where-object name -match "Office" | where-object
Caption -match "\d_"|%{$matches[0]} | select -first 1 #| ft -hide

# One more example
# $program1 = $win32Product | where-object name -match "SIMPLE PROGRAM NAME"
# if ($program1 -eq $null){$program1 = ""} else {$program1 = "SIMPLE PROGRAM NAME"}

$allProg = $check + ' ' + 'Office' + $office #+ ' ' + $program1

# Processor
$prc = get-ciminstance win32_processor|select-object -excludeProperty "CIM*" |
select Name | sort

# Monitor
$ManufacturerHash = @{
    "AAC" =	"AcerView";
    "ACR" = "Acer";
    "AOC" = "AOC";
    "AIC" = "AG Neovo";
    "APP" = "Apple Computer";
    "AST" = "AST Research";
    "AUO" = "Asus";
    "BNQ" = "BenQ";
    "CMO" = "Acer";
    "CPL" = "Compal";
    "CPQ" = "Compaq";
    "CPT" = "Chunghwa Pciture Tubes, Ltd.";
    "CTX" = "CTX";
    "DEC" = "DEC";
    "DEL" = "Dell";
    "DPC" = "Delta";
    "DWE" = "Daewoo";
    "EIZ" = "EIZO";
    "ELS" = "ELSA";
    "ENC" = "EIZO";
    "EPI" = "Envision";
    "FCM" = "Funai";
    "FUJ" = "Fujitsu";
    "FUS" = "Fujitsu-Siemens";
    "GSM" = "LG Electronics";
    "GWY" = "Gateway 2000";
    "HEI" = "Hyundai";
    "HIT" = "Hyundai";
    "HSL" = "Hansol";
    "HTC" = "Hitachi/Nissei";
    "HWP" = "HP";
    "IBM" = "IBM";
    "ICL" = "Fujitsu ICL";
    "IVM" = "Iiyama";
    "KDS" = "Korea Data Systems";
    "LEN" = "Lenovo";
    "LGD" = "Asus";
    "LPL" = "Fujitsu";
    "MAX" = "Belinea";
    "MEI" = "Panasonic";
    "MEL" = "Mitsubishi Electronics";
    "MS_" = "Panasonic";
    "NAN" = "Nanao";
    "NEC" = "NEC";
    "NOK" = "Nokia Data";
    "NVD" = "Fujitsu";
    "OPT" = "Optoma";
    "PHL" = "Philips";
    "REL" = "Relisys";
    "SAN" = "Samsung";
    "SAM" = "Samsung";
    "SBI" = "Smarttech";
    "SGI" = "SGI";
    "SNY" = "Sony";
    "SRC" = "Shamrock";
    "SUN" = "Sun Microsystems";
    "SEC" = "Hewlett-Packard";
    "TAT" = "Tatung";
    "TOS" = "Toshiba";
    "TSB" = "Toshiba";
    "VSC" = "ViewSonic";
    "ZCM" = "Zenith";
    "UNK" = "Unknown";
    "_YV" = "Fujitsu";
      }


  #Takes each computer specified and runs the following code:
  ForEach ($Computer in $computerName) {
    
    $Mon_Count = 1

    #Grabs the Monitor objects from WMI
    $Monitors = Get-WmiObject -Namespace "root\WMI" -Class "WMIMonitorID" -ComputerName $Computer -ErrorAction SilentlyContinue

    #Creates an empty array to hold the data
    $Monitor_Array = @()


    #Takes each monitor object found and runs the following code:
    ForEach ($Monitor in $Monitors) {

      #Grabs respective data and converts it from ASCII encoding and removes any trailing ASCII null values
      If ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName) -ne $null) {
        $Mon_Model = ([System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName)).Replace("$([char]0x0000)","")
      } else {
        $Mon_Model = $null
      }
      $Mon_Serial_Number = ([System.Text.Encoding]::ASCII.GetString($Monitor.SerialNumberID)).Replace("$([char]0x0000)","")
      $Mon_Attached_Computer = ($Monitor.PSComputerName).Replace("$([char]0x0000)","")
      $Mon_Manufacturer = ([System.Text.Encoding]::ASCII.GetString($Monitor.ManufacturerName)).Replace("$([char]0x0000)","")

      #Filters out "non monitors". Place any of your own filters here. These two are all-in-one computers with built in displays. I don't need the info from these.
      If ($Mon_Model -like "*800 AIO*" -or $Mon_Model -like "*8300 AiO*") {Break}

      #Sets a friendly name based on the hash table above. If no entry found sets it to the original 3 character code
      $Mon_Manufacturer_Friendly = $ManufacturerHash.$Mon_Manufacturer
      If ($Mon_Manufacturer_Friendly -eq $null) {
        $Mon_Manufacturer_Friendly = $Mon_Manufacturer
      }

      #Creates a custom monitor object and fills it with 4 NoteProperty members and the respective data
      $Monitor_Obj = [PSCustomObject]@{
        Title            = "Monitor ($Mon_Count):"
        Manufacturer     = $Mon_Manufacturer_Friendly
        Model            = $Mon_Model
        SerialNumber     = $Mon_Serial_Number
      }

          $Mon_Count = $Mon_Count + 1

      #Appends the object to the array
      $Monitor_Array += $Monitor_Obj

    } #End ForEach Monitor

    #Outputs the Array
    ($Monitor_Array |ft -Hide | where{$_ -ne ""} | Out-File -FilePath $savePath -Append | Format-List) 

} #End ForEach Computer
}

# Memory
$cap_mem = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb

# Hard disk
$hardDiskSizeBites = get-physicaldisk | sort Number | select size -ExpandProperty size
$hardDiskSize = ($hardDiskSizeBites[0] / 1gb).tostring("F00")

# PC model
$pcModel = gwmi Win32_ComputerSystem

# OS 
$osType = [string]::Format("{0} {1}", (Get-WmiObject Win32_OpetaingSystem).Caption, (wmic os get OSArchitecture)[2])

# Printers
$printers = Get-Printer |
Select Name,Location,PortName,Netip |
where{$_.Name -match "(\A[^\bMICROSOFT\b])"} |
where{$_.Name -match "[^FAX\b]\Z" |
where{$_.Name -match "(\A[^\bPDFCREATOR\b])"} |
where{$_.Name -match "(\A[^\bSEND\b])"} 

# Final. Information export
$skip | Out-File -FilePath $savePath -Append

'Full name: ' + $fullName | Out-File -FilePath $savePath -Append
'Programs: ' + $allProg | Out-File -FilePath $savePath -Append
'Processor: ' + $prc.Name | Out-File -FilePath $savePath -Append
Get-Monitors
'Memory: ' + $cap_mem + ' Gb' | Out-File -FilePath $savePath -Append
'Hard disk: ' + $hardDiskSize + ' Gb' | Out-File -FilePath $savePath -Append
'PC name: ' + $computerName | Out-FIle -FilePath $savePath -Append
'PC model: ' + $pcModel.Manufacture + $pcModel.Model | Out-File -FilePath $savePath -Append
'OS: ' + $osType | Out-File -FilePath $savePath -Append
$printers | ft -hide | Out-File -FilePath $savePath -Append

$skip | Out-File -FilePath $savePath -Append


# Bonus
# Delete empty strings
$nValue = 0
$ content = [System.IO.File]::ReadAllText($savePath)
$re = "(?m)(^\s*\r?\n){" + $nValue + ",}"
$res = $content -replace $re, ""
[System.IO.File]::WriteAllText($savePath, $res)

# Delete file (if you need delete label for example)
# $iconTest = Test-Path $savePath
# if ($iconTest -eq "True"){Remove-item ($iconPath)}else{return}

# Powershell create zombie-like process so ...
(Get-Process -Name powershell).Kill()
