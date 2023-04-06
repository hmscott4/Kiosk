# Hugh Scott
# press-key.ps1
# 2022/02/18
 
# Description: 
# 1. Start MS Edge with a set list of URIs.  Each URI will appear on a new tab.
# 2. Place MS Edge into Full Screen mode
# 3. Every 60 seconds, switch to a new tab
# 4. Every 5 minutes, refresh the browser
# 5. Every Sunday at 06:00, restart the computer
 
# Use with Kiosk GPO to limit user access to computer functions
# Use with auto logon enabled to allow user to automatically login
# Set as startup script (in GPO)

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet("Production","Test")]
    [string]
    $Environment="Production"
)

function Get-Config
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet("Production","Test")]
        [String]
        $Environment
    )

[xml]$config=@"
<?xml version="1.0" encoding="UTF-8"?>
<config>
    <environments>
        <environment name='Production'>
            <browser>
                <name>MicrosoftEdge_8wekyb3d8bbwe'</name>
                <path>"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"</path>
            </browser>
            <urls>
                <url>https://pbirs.abcd.lcl/reports/powerbi/Enterprise Dashboards/Infrastructure/Active Directory Domain Services</url>
                <url>https://pbirs.abcd.lcl/reports/powerbi/Enterprise Dashboards/Infrastructure/SQL Server</url>
                <url>https://pbirs.abcd.lcl/reports/powerbi/Enterprise Dashboards/Infrastructure/Hyper-V Dashboard</url>
            </urls>
            <rotation>
                <intervalSeconds>60</intervalSeconds>
                <refreshMinutes>300</refreshMinutes>
                <maxMinutes>10080</maxMinutes>
            </rotation>
            <reboot>
                <enabled>1</enabled>
                <day>Sun</day>
                <time>18:00</time>
            </reboot>
        </environment>
    </environments>
</config>
"@

$tmpConfig = $config.SelectSingleNode("//config/environments/environment[@name='$Environment']")
[xml]$retVal = "<config>" + $tmpConfig.InnerXml + "</config>"
return $retVal
}

################################################################################
# GET CONFIGURATION
# SET INITIAL VALUES
################################################################################
[xml]$thisConfig = Get-Config $Environment;
[bool]$fullScreen = $false;
$appName = $thisConfig.config.browser.name;
$appInvocation = $thisConfig.config.browser.path;
[string]$navTo = ""
foreach($url in $thisConfig.config.urls.url)
{
    $navTo = $navTo + [uri]::EscapeUriString($url) + " "
}
[int]$intervalSeconds = $thisConfig.config.rotation.intervalSeconds;
[int]$refreshMinutes = $thisConfig.config.rotation.refreshMinutes;
[int]$maxMinutes = $thisConfig.config.rotation.maxMinutes
[string]$rebootDay=$thisConfig.config.reboot.day
[string]$rebootTime=$thisConfig.config.reboot.time
$rebootEnabled=$thisConfig.config.reboot.enabled

################################################################################
# START MAIN PROCESS
################################################################################
$myShell = New-Object -ComObject "Wscript.Shell";
 
Start-Process -FilePath $appInvocation -ArgumentList $navTo
 
for($i=1; $i -lt $maxMinutes; $i++)
{
    if(!$fullScreen)
    {
       Start-Sleep -Seconds 10;
        $myShell.AppActivate("$appName") | out-null;
        $myShell.SendKeys("{F11}");
       $fullScreen = $true
    }
 
    ### Every $intervalSeconds, rotate to the next tab
    Start-Sleep -Seconds $intervalSeconds;
    $myShell.AppActivate("$appName") | out-null;
    $myShell.SendKeys("^{PGDN}");
 
    ### Every $RefreshMinutes, refresh the whole browser
    If(($i % $refreshMinutes) -eq 0)
    {
        $myShell.AppActivate("$appName") | out-null;
        $myShell.SendKeys("^{F5}");
    }

    ### At the specified day of week/time, reboot
    ### Bear in mind that if you increase intervalSeconds, 
    ### you could "skip over" the time specified
    ### Disable reboot by setting <enabled> to 0
    $dayString = (Get-Date).ToString("ddd")
    If (($dayString -eq $rebootDay) -and ($rebootEnabled -eq 1))
    {
        [datetime]$rebootNow = [datetime]::ParseExact($rebootTime, 'HH:mm',$null)
        [datetime]$currentTime = Get-Date
        If(($currentTime -gt $rebootNow) -and ($currentTime -lt $rebootNow.AddMinutes(5)))
        {
            Restart-Computer
        }
    }
} 
