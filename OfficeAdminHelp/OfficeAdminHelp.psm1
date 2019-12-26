<#
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

By installing, you have acknowledged the above disclaimer. Do not remove this section.
#>

#This lists all of the commands available at the current point in time
function View-Command
{
    Param([Switch]$Toolbox)
    if($Toolbox)
    {
        Write-Host "Clean-ProductKey `t`tRemoves Office 16 product keys from the local machine. `n`t`t`t`t-Status reviews status of installed key instead. `n`t`t`t`t-Key specifies a single 5 character key"
    }
    Else
    {
        Write-Host "Command`t`t`t`t Description`n-------------------------------------------------------------------------------"
        Write-Host "Install-AdminTool`t`t Installs Prerequisite modules"
        Write-Host "Update-AdminTool`t`t Updates Administration modules"
        Write-Host "Connect-Azure`t`t`t Connect to AzureAD"
        Write-Host "Connect-Exchange`t`t Connect to ExchangeOnline"
        Write-Host "Connect-MSOL`t`t`t Connect to MSOL Service"
        Write-Host "Connect-Security`t`t Connect to Security and Compliance center"
        Write-Host "Connect-SharePoint`t`t Connect to SharePoint Admin Center"
        Write-Host "Connect-SharePointPnP `t`t Connect to a specific SharePoint site"
        Write-Host "Connect-Skype`t`t`t Connect to Skype for Business Online"
        Write-Host "Connect-Multi`t`t`t Connect to multiple services at once"
        Write-Host "Save-Credentials`t`t Saves credentials to be used later"
        Write-Host "Logout-All`t`t`t Closes any active PSSession instances to avoid lockout"
        Write-Host "About-Command`t`t`t Provides additional information about Commands"
        Write-Host "View-Command -Toolbox `t`t Provides information about toolbox commands"
    }
}

#Provides additional information to learn more.
Function About-Command
{
    Param([Parameter(Mandatory)][String]$Command,[Switch] $Toolbox)
    if(-not($Toolbox))
    {
        if($Command -eq "Install-AdminTool")
        {
            Write-Host "Installs Prerequisite modules for the Admin tool to function.`n`nThis includes the following modules from Microsoft:`n"
            Write-Host "AzureAD`nMSOnline`nMicrosoft.Online.SharePoint.PowerShell`nSharePointPnPPowerShellOnline`nSkypeOnlineConnector"
            Write-Host "`nAdditional dependencies can be installed from the following locations`nhttps://www.microsoft.com/en-us/download/details.aspx?id=39366`nhttps://www.microsoft.com/en-us/download/details.aspx?id=41950`n"
        }
        if($Command -eq "Update-AdminTool")
        {
            Write-Host "Updates Prerequisite modules acquired from psgallery for the Admin tool to function.`n`nThis includes the following modules from Microsoft:`n"
            Write-Host "AzureAD`nMSOnline`nMicrosoft.Online.SharePoint.PowerShell`nSharePointPnPPowerShellOnline`n"
        }    
        if($Command -eq "Connect-Azure")
        {
            Write-Host "Creates a connection to azure AD for the use of CMDLETS found here: `nhttps://docs.microsoft.com/en-us/powershell/module/azuread/?view=azureadps-2.0"
        }
        if($Command -eq "Connect-Exchange")
        {
            Write-Host "Creates a connection to Exchange Online  for the use of CMDLETS found here in the reference section: `nhttps://docs.microsoft.com/en-us/powershell/exchange/exchange-online/exchange-online-powershell?view=exchange-ps"
        }
        if($Command -eq "Connect-MSOL")
        {
            Write-Host "Creates a connection to Microsoft Online Services for the use of CMDLETS found here: `nhttps://docs.microsoft.com/en-us/powershell/module/msonline/?view=azureadps-1.0"
        }
        if($Command -eq "Connect-Security")
        {
            Write-Host "Creates a connection to the Security and Compliance Center for the use of CMDLETS found here: `nhttps://docs.microsoft.com/en-us/microsoft-365/compliance/labels#find-the-powershell-cmdlets-for-labels"
        }
        if($Command -eq "Connect-Sharepoint")
        {
            Write-Host "Creates a connection to SharePoint's admin utility for the use of CMDLETS found here: `nhttps://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/introduction-sharepoint-online-management-shell?view=sharepoint-ps"
        }
        if($Command -eq "Connect-SharepointPNP")
        {
            Write-Host "Creates a connection to a specific SharePoint or OneDrive site for the use of CMDLETS found here: `nhttps://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps"
        }
        if($Command -eq "Connect-Skype")
        {
            Write-Host "Creates a connection to Skype for Business Online for the use of CMDLETS found here: `nhttps://docs.microsoft.com/en-us/powershell/module/skype/?view=skype-ps"
        }

    }
    if($Toolbox)
    {
        if($Command -eq "Clean-ProductKey")
        {
            Write-host "Cleans product keys from office 365 from the local machine. `nThe following switches can be run in conjuntion with this command:`n -status creates a file at the root of C:/ to review later to ensure no keys are expired `n-key and the last 5 digits of a product key can be used to clean a specified expired key off of the local machine`n there is no -source switch available for this command. Review the module file itself."
        }
        if($Command -eq "Get-Groupmembers")
        {
            Write-Host "Gathers information about distribution group members given. Can be called with -asfile to return results as a CSV located at C:/ root. `nUse -source to give the source of the asfile prompt. remove the final piped command to return information to the console "
        }
    }
}