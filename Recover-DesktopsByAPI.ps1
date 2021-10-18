#requires -Version 3.0
$ErrorActionPreference = 'SilentlyContinue'
<#
    .SYNOPSIS
    Recovers a VMware Horizon VDI machine

    .DESCRIPTION
    This script Recovers a VMware Horizon Instant Clone VDI machine using the VMware Horizon API's

    .EXAMPLE
    You can use this to 'rebuild' an Instant Clone if there is an issue with the machine.

    .NOTES
    This script requires VMWare PowerCLI to be installed on the machine running the script.
    PowerCLI can be installed through PowerShell (PowerShell version 5 or higher required) by running the command 'Install-Module VMWare.PowerCLI -Force -AllowCLobber -Scope AllUsers'

    Before running this script you will also need to have a PSCredential object available on the target machine. This can be created by running the 'Create credentials for VMware Horizon scripts' script in ControlUp on the target machine.
    
    Other prerequisites:

    Set-PowerCLIConfiguration -Scope AllUsers -ParticipateInCEIP $false
    Install powercli modules offline and put them in modules path:     https://developer.vmware.com/powercli/installation-guide
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope allusers -Confirm:$false

    Context: Can be triggered from the VMware Horizon Machines view
    Modification history:   09/01/2020 - Anthonie de Vreede - First version
                            18/05/2021 - Wouter Kursten - Second version
                            15/10/2021 - Bart Tacken - Edited to work without ControlUP https://www.controlup.com/script-library-posts/horizon-recover-instant-clone-vdi-machine/

    Changelog:
        - 18/05/2021: remove everything related to the VMware.hv.helper

    .PARAMETER strHVMachineName
    Name of the VMware Horizon machine. Passed from the ControlUp Console.
    .PARAMETER strHVMachinePool
    Name of the VMware Horizon machine Pool. Passed from the ControlUp Console.
    .PARAMETER strHVMachineSource
    Type of machine. Passed from ControlUp Console.
    .PARAMETER strHVConnectionServerFQDN
    Name of the VMware Horizon connection server. Passed from the ControlUp Console.
    
    .LINK
    https://code.vmware.com/web/tool/11.3.0/vmware-powercli

    
    .COMPONENT
    VMWare PowerCLI 11 or higher
#>

$strHVMachinePool = "FL-W10-Basic-TEST"
$strHVMachineSource = "vCenter (instant clone)"
$strHVConnectionServerFQDN = ""

$KeyFilePath = 'C:\Scripts\VDIscript\customer_.key'
$CredPath = 'C:\Scripts\VDIscript\customer_Cred.xml'

Function Out-CUConsole {
    <# This function provides feedback in the console on errors or progress, and aborts if error has occured.
      If only Message is passed this message is displayed
      If Warning is specified the message is displayed in the warning stream (Message must be included)
      If Stop is specified the stop message is displayed in the warning stream and an exception with the Stop message is thrown (Message must be included)
      If an Exception is passed a warning is displayed and the exception is thrown
      If an Exception AND Message is passed the Message message is displayed in the warning stream and the exception is thrown
    #>

    Param (
        [Parameter(Mandatory = $false)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )

    # Throw error, include $Exception details if they exist
    if ($Exception) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        If ($Message) {
            Write-Warning -Message "$Message`n$($Exception.CategoryInfo.Category)`nPlease see the Error tab for the exception details."
            Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)`n$($Exception.Exception.ErrorRecord)" -ErrorAction Stop
        }
        Else {
            Write-Warning "There was an unexpected error: $($Exception.CategoryInfo.Category)`nPlease see the Error tab for details."
            Throw $Exception
        }
    }
    elseif ($Stop) {
        # Write simplified error message to Warning stream, Throw exception with simplified message as well
        Write-Warning -Message "There was an error.`n$Message"
        Throw $Message
    }
    elseif ($Warning) {
        # Write the warning to Warning stream, thats it. It's a warning.
        Write-Warning -Message $Message
    }
    else {
        # Not an exception or a warning, output the message
        Write-Output -InputObject $Message
    }
}
function Load-VMWareModules {
    <# Imports VMware modules
    NOTES:
    - The required modules to be loaded are passed as an array.
    - In versions of PowerCLI below 6.5 some of the modules can't be imported (below version 6 it is Snapins only) using so Add-PSSnapin is used (which automatically loads all VMWare modules)
    #>

    param (    
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware module to be loaded. Can be single or multiple values (as array).")]
        [array]$Components
    )

    # Try Import-Module for each passed component, try Add-PSSnapin if this fails (only if -Prefix was not specified)
    # Import each module, if Import-Module fails try Add-PSSnapin
    foreach ($component in $Components) {
        try {
            $null = Import-Module -Name VMware.$component
        }
        catch {
            try {
                $null = Add-PSSnapin -Name VMware
            }
            catch {
                #Out-CUConsole -Message 'The required VMWare modules were not found as modules or snapins. Please check the .NOTES and .COMPONENTS sections in the Comments of this script for details.' -Stop
            }
        }
    }
}
function Connect-HorizonConnectionServer {
    param (    
        [parameter(Mandatory = $true,
            HelpMessage = "The FQDN of the VMware Horizon Connection server. IP address may be used.")]
        [string]$HVConnectionServerFQDN,
        [parameter(Mandatory = $true,
            HelpMessage = "The PSCredential object used for authentication.")]
        [PSCredential]$Credential
    )

    try {
        Write-Host "Making connecting with Horizon server.."
        Connect-HVServer -Server $HVConnectionServerFQDN -Credential $Credential -ErrorAction SilentlyContinue
    }
    catch {
        if ($_.Exception.Message.StartsWith('Could not establish trust relationship for the SSL/TLS secure channel with authority')) {
            Out-CUConsole -Message 'There was a problem connecting to the VMware Horizon Connection server. It looks like there may be a certificate issue. Please ensure the certificate used on the VMware Horizon server is trusted by the machine running this script.' -Exception $_
        }
        else {
            Out-CUConsole -Message 'There was a problem connecting to the VMware Horizon Connection server.' -Exception $_
        }
    }
}
function Disconnect-HorizonConnectionServer {
    param (    
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    
    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false
    }
    catch {
        Out-CUConsole -Message 'There was a problem disconnecting from the VMware Horizon Connection server. If not running in a persistent session (ControlUp scripts do not run in a persistant session) this is not a problem, the session will eventually be deleted by VMware Horizon.' -Warning
    }
}
function Recover-HorizonViewMachine {
    param (   
        [parameter(Mandatory = $true,
            HelpMessage = "The VMware Horizon machine object.")]
        [object]$HVMachineid,
        [parameter(Mandatory = $true,
        HelpMessage = "The VMware Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        $HVConnectionServer.extensiondata.Machine.Machine_Recover($HVMachineId)
        Out-CUConsole -Message 'Recover command has been sent to VMware Horizon.'
    }
    catch {
        Out-CUConsole -Message 'There was a problem Recovering the machine.' -Exception $_
    }
}
function Get-HVDesktopPool {
    param (
        [parameter(Mandatory = $true,
        HelpMessage = "Displayname of the Desktop Pool.")]
        [string]$HVPoolName,
        [parameter(Mandatory = $true,
        HelpMessage = "The VMware Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )
    # Try to get the Desktop pools in this pod
    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'DesktopSummaryView'
        # Filter on the correct displayname
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='desktopSummaryData.displayName'; 'value' = "$HVPoolname"}
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults){
            Out-CUConsole -Message "Can't find $HVPoolName, exiting."
            exit
        }
        else {
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the VMware Horizon Desktop Pool.' -Exception $_
    }
}
function Get-HVDesktopMachine {
    param (
        [parameter(Mandatory = $true,
        HelpMessage = "ID of the Desktop Pool.")]
        [VMware.Hv.DesktopId]$HVPoolID,
        [parameter(Mandatory = $true,
        HelpMessage = "Name of the Desktop machine.")]
        [string]$HVMachineName,
        [parameter(Mandatory = $true,
        HelpMessage = "The VMware Horizon Connection server object.")]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        # create the service object first
        [VMware.Hv.QueryServiceService]$queryService = New-Object VMware.Hv.QueryServiceService
        # Create the object with the definiton of what to query
        [VMware.Hv.QueryDefinition]$defn = New-Object VMware.Hv.QueryDefinition
        # entity type to query
        $defn.queryEntityType = 'MachineDetailsView'
        # Filter so we get the correct machine in the correct pool
        $poolfilter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='desktopData.id'; 'value' = $HVPoolID}
        $machinefilter = New-Object VMware.Hv.QueryFilterEquals -property @{'memberName'='data.name'; 'value' = "$HVMachineName"}
        $filterlist = @()
        $filterlist += $poolfilter
        $filterlist += $machinefilter
        $filterAnd = New-Object VMware.Hv.QueryFilterAnd
        $filterAnd.Filters = $filterlist
        $defn.Filter = $filterAnd
        # Perform the actual query
        [array]$queryResults= ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        # Remove the query
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)
        # Return the results
        if (!$queryResults){
            Out-CUConsole -Message "Can't find $HVPoolName, exiting."
            exit
        }
        else{
            return $queryResults
        }
    }
    catch {
        Out-CUConsole -Message 'There was a problem retreiving the VMware Horizon Desktop Pool.' -Exception $_
    }
}

# Check if the machine is an Instant Clone. Recovery only works for Instant and Linked Clones
if ($strHVMachineSource -ne 'vCenter (instant clone)') {
    Out-CUConsole -Message 'This machine is not an Instant Clone, it cannot be recoverred.' -Stop
}
[string]$DateStr = (Get-Date).ToString("s").Replace(":","-") # +"_" # Easy sortable date string  
Start-Transcript -Path ('C:\Scripts\VDIscript\'  + $DateStr  + 'Recover-DesktopsByAPI.log') -Force

Write-Host "Import the VMware PowerCLI modules"
#Load-VMwareModules -Components @('VimAutomation.HorizonView')
#Write-Host "Loading VMware modules" -ForegroundColor Green
Import-Module -Name VMware.VimAutomation.HorizonView -Force
#Import-Module -Name VMware*

Write-Host "Succesfully loaded PowerCLI modules"
#region Get the stored credentials for running the script


#decrypt credentials
$Key = Get-Content $KeyFilePath
$credXML = Import-Clixml $CredPath # Import encrypted credential file into XML format
$secureStringPWD = ConvertTo-SecureString -String $credXML.Password -Key $key
$CredsHorizon = New-Object System.Management.Automation.PsCredential($credXML.UserName, $secureStringPWD) # Create PScredential Object
#endregion

# Connect to the VMware Horizon Connection Server
[VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $strHVConnectionServerFQDN -Credential $CredsHorizon

$VMs = (Get-HVMachine -PoolName $strHVMachinePool).base.name

foreach ($VM in $VMs) {

    Write-Host "Recover VM [$VM]"  
    [Vmware.Hv.machineid]$machine=(get-Hvmachine -machinename $VM).id

    # Recover the VMware Horizon Machine
    Recover-HorizonViewMachine -HVMachineid $machine -HVConnectionServer $objHVConnectionServer
}

# Disconnect from the VMware Horizon Connection Center
Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
Stop-Transcript
