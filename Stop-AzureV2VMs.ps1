<# 
    .SYNOPSIS 
        This Azure Automation runbook automates the scheduled shutdown and startup of virtual machines in an Azure subscription.
        Author:  Gregory Van Den Ham
        Date:  12 December 2017
		This code block is the shutdown script.
		
    .DESCRIPTION 
        The runbook implements a solution for scheduled power management of Azure virtual machines.  It accepts VMName, ResourceGroup or data Stored
        in the PowerStatus Tag.  If no information is provided, this runbook will manage the power status of all VM's in the subscription.
 
        This is a PowerShell runbook, as opposed to a PowerShell Workflow runbook. 
 
        This runbook requires the "Azure" and "AzureRM.Resources" modules which are present by default in Azure Automation accounts. 
 
    .PARAMETER VMName 
        The name of the Virtual Machine to Power Manage.

    .PARAMETER ResourceGroupName 
        The name of the Resource Group to Power Manage.

    .PARAMETER PowerStatusTag 
        The Tag data in the Tag Named PowerStatus to Power Manage.  This script currently only accepts PowerStatus : AzureAutomatedShutdownStartup.

#>

param (
  
    [Parameter(Mandatory=$false,HelpMessage="Provide the name of the VM, if none we'll look at the resourcegroup name")] 
    [String] $VMName ,
 
    [Parameter(Mandatory=$false,HelpMessage="Provide the name of the ResourceGroup, all systems in group will be impacted")] 
    [String] $ResourceGroupName ,

    [Parameter(Mandatory=$false,HelpMessage="Provide the contents of the PowerStatusTag, all systems with that tag will be impacted")] 
    [String] $PowerStatusTag    
)

# What Parameters were passed    
Write-Output ("Processing variables VMName= " + $VMName  + " , ResourceGroupName= " + $ResourceGroupName + "PowerStatus Tag= " + $PowerStatusTag + ".")

$connectionName = "AzureRunAsConnection"
try
{
    # Get the connection "AzureRunAsConnection "
    $servicePrincipalConnection=Get-AutomationConnection -Name $connectionName         
 
    "Logging in to Azure..."
    Add-AzureRmAccount `
        -ServicePrincipal `
        -TenantId $servicePrincipalConnection.TenantId `
        -ApplicationId $servicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint 
}
catch {
    if (!$servicePrincipalConnection)
    {
        $ErrorMessage = "Connection $connectionName not found."
        throw $ErrorMessage
    } else{
        Write-Error -Message $_.Exception
        throw $_.Exception
    }
}
 
### Fixed Logig 8 Dec 2017 Greg V. 
# If there is a specific resource group and vmname grab that specific vm.
# If there is a specific resource group, then get all VMs in the resource group.
# If there is just a VMName grab just the VM.
# If there is a proper tag read the value and try to use it.
# IF VM and Resource is blank otherwise get all VMs in the subscription.

######
# Stop Specific VM in Specific ResourceGroup
if (($ResourceGroupName -ne '') -And ($VMName -ne '')) 
{ 
    $VMs = Get-AzureRmVM -ResourceGroupName $ResourceGroupName -Name $VMName
    Write-Output("Case 1 - Using ResourceName and VMName to target specific VM for shutdown.")
}
#####
# Stop Entire ResourceGroup
elseif (($ResourceGroupName -ne '') -And ($VMName -eq ''))
{
      $VMs = Get-AzureRmVM -ResourceGroupName $ResourceGroupName  
      Write-Output("Case 2 - Using ResourceName - targeting all VM's in ResourceGroup for shutdown.")
}
#####
# Stop Specific VM but don't know its Resource Group
elseif (($ResourceGroupName -eq '') -And ($VMName -ne ''))
{
    $AllVMs = Get-AzureRmVM
    foreach ($ListedVM in $AllVMs)
    {
        if ($ListedVM.Name -eq $VMName)
        {
            $VMResGroup = $ListedVM.ResourceGroupName
            Write-Output("Looked up Resource group for: " + $VMName + ". Found in " + $VMResGroup)
            $VMs = Get-AzureRMVM -ResourceGroupName $VMResGroup -Name $VMName
        }    
    }
    Write-Output("Case 3 - Using VMName only - Took time to find VM's Resource Group in Subscription to be able to Target VM for shutdown.")
}
######
# Stop Based on information on Tag PowerStatus  ** Currently basic, but can expand
elseif (($ResourceGroupName -eq '') -And ($VMName -eq '') -And ($PowerStatusTag -ne ''))
{
$VMs = @()
$AllVMs = Get-AzureRmVM
    foreach ($ListedVM in $AllVMs)
    {
        if($ListedVM.Tags.Keys -contains "PowerStatus")
        {
            $tag_value=$ListedVM.Tags.PowerStatus
            write-host ($ListedVM.Name + " has requested power status: " + $tag_value)
                if ($tag_value -eq  "AzureAutomatedShutdownStartup")
                {
                $VMResGroup = $ListedVM.ResourceGroupName
                Write-Output("Looked up Resource group for: " + $ListedVM.Name + ". Found in " + $VMResGroup)
                $VMInfodetail = Get-AzureRMVM -ResourceGroupName $VMResGroup -Name $ListedVM.Name
                
                $row = new-object PSObject -Property @{
                    Name = $VMInfodetail.Name;
                    ResourceGroupName = $VMInfodetail.ResourceGroupName
                    StatusCode = $VMInfodetail.StatusCode
                    }
                    $VMs += $row
                    
                }
         }
     } 
    Write-Output("Case 4 - Using PowerStatus Tag - targeting all VM's with properly definded PowerStatus Tag for shutdown.")
}
#####
# Stop Everything
else 
{ 
    $VMs = Get-AzureRmVM
    Write-Output("Case 5 - No ResourceName or VMName given - All VM's in Subscription will be attempted to shutdown.")
}

# Logic added - Display VM's we're about to stop 8 Dec 2017 GregV
# Write what we're about to stop
foreach ($VM in $VMs)
{
    $VMShutdownList += $VM.Name + " "
}
Write-Output ("Hang on we're queueing " + $VMShutdownList + " for shutdown.") 

####Shutdown VMs ####

# Shutdown VMs - Rewritten 11Dec2017 GregV
if(!$VMs) 
    {
    Write-Output ("No VMs were found with the parameters specified to shutdown.")
    }
    else 
    {
        ForEach ($VM in $VMs) 
        {
            try
            {
                $StopVM = Stop-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name -ErrorAction Continue -Force
            }
            catch
            {
                Write-Error ($VM.Name + " on resource group "+ $VM.ResourceGroupName + " failed to stop. Error was: ") -ErrorAction Continue
                Write-Error ("Status was "+ $StopVM.Status) -ErrorAction Continue
                Write-Error (ConvertTo-Json $StopVM.Error) -ErrorAction Continue
                $errorlinenumber = $_.InvocationInfo.ScriptLineNumber
                Write-Error ("Error Occurred on Line: " + $errorlinenumber + ".")
            }
            $Attempt = 1
            if(($StopVM.StatusCode) -ne 'OK')
            {
                do
                {
                    Write-Output ("Failed to stop " + $VM.Name +" . Retrying in 60 seconds...")
                    write-Output ("Current VM Status is: " + $StopVM.StatusCode + ".")
                    $VMRunstatus=(Get-AzureRmVM $VM.Name -ResourceGroupName $VM.ResourceGroupName -Status).Statuses | where Code -like "PowerState*"
                    Write-Output ("Current Condition is: " + $VMRunStatus.DisplayStatus + ".")   #Expected Conditions "VM deallocated" or "VM running"
                    Start-Sleep -Seconds 60
                    try
                    {
                        $StopVM = Stop-AzureRmVM -ResourceGroupName $VM.ResourceGroupName -Name $VM.Name -ErrorAction Continue -Force
                    }
                    catch
                    {
                        Write-Error ($VM.Name + " on resource group "+ $VM.ResourceGroupName + " failed to stop. Error was: ") -ErrorAction Continue
                        Write-Error ("Status was "+ $StopVM.Status) -ErrorAction Continue
                        Write-Error (ConvertTo-Json $StopVM.Error) -ErrorAction Continue
                        $errorlinenumber = $_.InvocationInfo.ScriptLineNumber
                        Write-Error ("Error Occurred on Line: " + $errorlinenumber + ".")                    
                    }
                    $Attempt++
                }
                while(($StopVM.StatusCode) -ne 'OK' -and $Attempt -lt 3)
            }
           
            if($StopVM)
            {
                Write-Output ("Stop-AzureRmVM cmdlet for: " + $VM.Name + " with StatusCode: " + $StopVM.StatusCode + " on attempt number: " + $Attempt +  " of 3.")
            }
        $VMRunstatus=(Get-AzureRmVM $VM.Name -ResourceGroupName $VM.ResourceGroupName -Status).Statuses | where Code -like "PowerState*"
        Write-Output ("Stop was attempted on: " + $VM.Name + ". Current Condition is: " + $VMRunStatus.DisplayStatus + ".")
        } 
    }