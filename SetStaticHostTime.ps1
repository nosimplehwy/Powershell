# ---------------------------------------------------------
# Script to setup static ip, hostname and time on devices that are defined in a spreadsheet
# Performs autodiscovery and sets up devices that match by mac address defined in the spreadsheet
# ---------------------------------------------------------

#Set-ExecutionPolicy RemoteSigned
Import-Module PSCrestron


# ---------------------------------------------------------
function Set-NetworkSettings
# ---------------------------------------------------------
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory=$true)]
		[object]$CurrentIP,

		[Parameter(Mandatory=$true)]
		[object]$NewSettings
    )

    try
    {
	    # open a socket
	    $s = Open-CrestronSession $CurrentIP -Secure -Username $user -Password $pass
		
        # set the IP address
	    Invoke-CrestronSession $s "IPA 0 $($NewSettings.IPA)" | Out-Null

        # set the IP mask
	    Invoke-CrestronSession $s "IPM 0 $($NewSettings.IPM)" | Out-Null

        # set the default router
	    Invoke-CrestronSession $s "DEFR 0 $($NewSettings.Gateway)" | Out-Null

        # set the DNS server
	    Invoke-CrestronSession $s "ADDD $($NewSettings.DNS1)" | Out-Null
	    Invoke-CrestronSession $s "ADDD $($NewSettings.DNS2)" | Out-Null

        # set the domain
	    #Invoke-CrestronSession $s "DOMAIN $($Static.Domain)" | Out-Null

        # set DHCP off
	    Invoke-CrestronSession $s 'DHCP 0 OFF' | Out-Null

        # get default hostname
        $hostname = Invoke-CrestronSession $s "HOST"

        # set the hostname
	    Invoke-CrestronSession $s "HOST $($NewSettings.Hostname + $hostname)" | Out-Null

	    # close the socket
	    Close-CrestronSession $s

        # reset the device but don't wait for it to recover
        Reset-CrestronDevice $CurrentIP -NoWait -Secure -Username $user -Password $pass | Out-Null
       
    }
    catch
        {
            Write-Host "An error occurred."
            throw
        }
}

# ---------------------------------------------------------
function Set-Time
# ---------------------------------------------------------
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory=$true)]
		[object]$CurrentIP

    )

    try
    {
	    # open a socket
	    $s = Open-CrestronSession $CurrentIP -Secure -Username $user -Password $pass
		
        # set the date/time to the computer time
        $time = Get-Date -Format "hh:mm:ss MM-dd-yyyy"
	    Invoke-CrestronSession $s "Time $time" | Out-Null

        #set the timezone
	    Invoke-CrestronSession $s "Timezone 7" | Out-Null

	    # close the socket
	    Close-CrestronSession $s       
    }
    catch
        {
            Write-Host "An error occurred."
            throw
        }
}

# ---------------------------------------------------------
function Set-AutoUpdate
# ---------------------------------------------------------
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory=$true)]
		[object]$CurrentIP
    )

    try
    {
	    # open a socket
	    $s = Open-CrestronSession $CurrentIP -Secure -Username $user -Password $pass
		
        #disable autoupdate
        Invoke-CrestronSession $s "auenable off" | Out-Null

	    # close the socket
	    Close-CrestronSession $s       
    }
    catch
        {
            Write-Host "An error occurred."
            throw
        }
}

# ---------------------------------------------------------
function Get-Status
# ---------------------------------------------------------
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory=$true)]
		[object]$CurrentIP
    )

}
# ---------------------------------------------------------
# main program starts here
# ---------------------------------------------------------

Write-Host "Set static ip address, hostname, time, timezone and disable auto update"
Write-Host "Select excel spreadsheet with device information:"
# import the spreadhseet
#Begin with File Select
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $false # Multiple files cannot be chosen
	#Filter = '*.xlsx | *.csv' # Specified file types
}
 
[void]$FileBrowser.ShowDialog()


If($null -eq $FileBrowser)
{
# error message and exit if no file
Write-Host "No file was loaded"
Read-Host -Prompt 'Press enter to exit'
Exit
}

$wbpath = $FileBrowser.FileName;

# Define variables from user input
$wsname = Read-Host -Prompt 'Enter worksheet name:'


Write-Host 'Importing the spreadsheet...'
try{
$ws = Import-Excel -Workbook $wbpath -Worksheet $wsname | Where-Object MACaddress 
}
catch
{
    Write-Host "Error loading spreadsheet."
    Exit
}

Write-Host 'Enter device login credentials...'

$credentials = Get-Credential
$user = $credentials.UserName
$pass = $credentials.GetNetworkCredential().Password

Write-Host "Starting update at $(Get-Date -Format G)..."

# auto discover the devices in the subnet, and get version info on all devices that return a valid mac address
Write-Host 'Running Auto-Discovery...'
$devs = Get-AutoDiscovery |
    Select-Object -ExpandProperty IP |
    Get-VersionInfo -Secure -Username $user -Password $pass |
    Where-Object MACAddress -Match '[A-F\d\.]+' 

# error message and exit if no devices found
if(!$devs) 
{
    Write-Host 'No devices found.'
    Read-Host -Prompt 'Press enter to exit'
    Exit
}

# iterate the MAC addresses
Write-Host 'Running each row...'

#create an object to store devices that were found and updated
$DeviceTable = @()


    $ws | Foreach-Object{
            try {
                
                $dev = $devs | Where-Object MACAddress -eq $_.MACAddress 

                if($dev){
                    Write-Host "Setting up $($_.Hostname)"
                    Set-Time -CurrentIP $dev.IPAddress -NewSettings $_
                    Set-AutoUpdate -CurrentIP $dev.IPAddress -NewSettings $_
                    # Set-NetworkSettings -CurrentIP $dev.IPAddress -NewSettings $_
                    $DeviceTable += $dev.IPAddress;
                }
                else {
                    Write-Host "No matching device was found for $($_.Hostname)"            
                }   
            }
            catch {
                Write-Output $_
            }
        }

#create an object to store status info of devices that were found and updated
$StatusTable = @()

    # for each device in the table, report status
    Write-Host 'Checking devices...'
   foreach($device in $DeviceTable)
    {
        try
        {                  

               # run get-versioninfo and collect the info we want 
               $StatusRow =  Get-VersionInfo -Device $device -Secure -Username $user -Password $pass | Select-Object -Property Hostname,MACAddress

               #open a socket
               $Session = Open-CrestronSession -Device $device -Secure -Username $user -Password $pass
                $x = 0;
                #run ipconfig command and collect the info we want and add it to the $deviceinfo object
                $ipinfo = Invoke-CrestronSession -Handle $Session -Command 'ipconfig' -ErrorAction SilentlyContinue 
                $ipinfo -split "`n" | Select-String -Pattern '.+ (\d+\.\d+.\d+.\d+)' -AllMatches | 
                ForEach-Object { $data = $_.Matches.Groups[1].Value
                                       $x++
                                       Switch($x){
                                       1{ Add-Member -InputObject $StatusRow -NotePropertyName 'IPA' -NotePropertyValue $data}
                                       2{ Add-Member -InputObject $StatusRow -NotePropertyName 'IPM' -NotePropertyValue $data}
                                       3{ Add-Member -InputObject $StatusRow -NotePropertyName 'Gateway' -NotePropertyValue $data}
                                       4{ Add-Member -InputObject $StatusRow -NotePropertyName 'DNS1' -NotePropertyValue $data}
                                       5{ Add-Member -InputObject $StatusRow -NotePropertyName 'DNS2' -NotePropertyValue $data}
                                       }
                                    }

                $time = Invoke-CrestronSession -Handle $Session -Command 'time' -ErrorAction SilentlyContinue |
                Select-String -Pattern ".+ (\d\d:\d\d:\d\d\s\d\d-\d\d-\d{4})" 
                Add-Member -InputObject $StatusRow -NotePropertyName 'Time' -NotePropertyValue  $time.Matches.Groups[1].Value  

                $timezone = Invoke-CrestronSession -Handle $Session -Command 'timezone' -ErrorAction SilentlyContinue |
                Select-String -Pattern '.+ \((.+)\)'
                Add-Member -InputObject $StatusRow -NotePropertyName 'TimeZone' -NotePropertyValue  $timezone.Matches.Groups[1].Value
                
                $autoupdate = Invoke-CrestronSession -Handle $Session -Command 'auenable' -ErrorAction SilentlyContinue |
                Select-String -Pattern '.+: (\w+)'
                Add-Member -InputObject $StatusRow -NotePropertyName 'AutoUpdate' -NotePropertyValue  $autoupdate.Matches.Groups[1].Value

                   Close-CrestronSession -Handle $Session

                   $StatusTable += $StatusRow
        }
        catch
            {
                Write-Host "Error"
                Write-Warning $_.Exception.GetBaseException().Message
            }
    }

    # send the results to Out-Gridview
    $StatusTable | Out-GridView

#>
# complete message
Write-Host "Completed at $(Get-Date -Format G)."
Read-Host -Prompt 'Press enter to exit'

