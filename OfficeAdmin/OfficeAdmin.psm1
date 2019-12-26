<#
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

By installing, you have acknowledged the above disclaimer. Do not remove this section.
#>

#This function provides connectivity via saved credentials to Azure AD
function Connect-Azure
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    If($Source)
    {
        Write-Host "The following code can be called to connect to Azure AD`n`$Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'`nConnect-AzureAD -Credential `$Credential"
    }
    else
    {
        Write-Host "Connecting to Azure AD using saved credentials"
        $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
        Connect-AzureAD -Credential $Credential
    }
}

#This function provides connectivity via saved credentials to Exchange Online
function Connect-Exchange 
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    if($Source)
    {
        Write-Host "The following code can be called to connect to Exchange Online`n`$Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'`n`$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential  `$Credential -Authentication Basic -AllowRedirection`nImport-PSSession `$Session -DisableNameChecking"
    }
    Else
    {
         Write-Host "Connecting to Exchange Online using saved credentials"
         $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
         $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential  $Credential -Authentication Basic -AllowRedirection
         Start-Sleep -Seconds 5
         Import-Module(Import-PSSession $Session -DisableNameChecking) -Global
    }
}

#This function provides connectivity via saved credentials to Microsoft Services Online
function Connect-MSOL
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    If($Source)
    {
        Write-Host "The following code can be called to connect to MSOL Service`n`$Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'`nConnect-MsolService -Credential `$Credential"
    }
    Else
    {
        Write-Host "Connecting to MSOL Service using saved credentials"
        $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
        Connect-MsolService -Credential $Credential
    }
}

#This function provides connectivity via saved credentials to Office 365's security and compliance center
function Connect-Security
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    If($Source)
    {
        Write-Host "The following code can be called to connect to the Security and Compliance Center `$Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'`n`$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential `$Credential -Authentication Basic -AllowRedirection`nImport-PSSession `$Session -DisableNameChecking"
    }
    Else
    {
        Write-Host "Connecting to Security and Compliance Center using saved credentials"
        $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
        Import-Module(Import-PSSession $Session -DisableNameChecking) -Global
    }
}

#This function provides connectivity via saved credentials to Sharepoint's Admin Center
function Connect-SharePoint
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    If ($Source)
    {
        Write-Host "The following code can be called to connect to SharePoint Admin Center:`n`$Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'`n`$orgName= Read-Host "Please Enter the org name EG: Contoso"`nConnect-SPOService -Url https://`$orgName-admin.sharepoint.com -Credential `$Credential"
    }
    Else
    {
        Write-Host "Connecting to SharePoint Admin Center using saved credentials"
        $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
        $orgName= Read-Host "Please Enter the org name EG: Contoso"
        Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $Credential
    }
}

#This function provides connectivity via saved credentials to a specific Sharepoint or OneDrive site via the Patterns and Practices Powershell
function Connect-SharePointPnP
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    If ($Source)
    {
        Write-Host "The following code can be called to connect to a specific SharePoint site using Patterns and Practices PowerShell`$Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'`nUpdate-Module SharePointPnPPowerShell*`$URL=Read-Host "Please enter the URL of the sharepoint/onedrive site"`nConnect-PnPOnline -Url `$URL –Credentials `$Credential" 
    }
    Else
    {
        Write-Host "Connecting to SharePoint Site using saved credentials"
        $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
        Update-Module SharePointPnPPowerShell*
        $URL=Read-Host "Please enter the URL of the sharepoint/onedrive site"
        Connect-PnPOnline -Url $URL –Credentials $Credential
    }
}

#This function provides connectivity via saved credentials to Skype for Business Online
function Connect-Skype
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    If ($Source)
    {
        Write-Host "The following code can be called to connect to Skype For Business Online`n`$Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'`nImport-Module SkypeOnlineConnector`n`$sfbSession = New-CsOnlineSession -Credential `$Credential`nImport-PSSession `$sfbSession"
    }
    Else
    {
        Write-Host "Connecting to Skype for Business Online using saved credentials"
        $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
        Import-Module SkypeOnlineConnector
        $sfbSession = New-CsOnlineSession -Credential $Credential
        Import-Module(Import-PSSession $sfbSession) -Global
    }
}

#This function provides connectivity to multiple services at once by calling the respective function from above.
function Connect-Multi
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'))
    {
        Write-Error -Message "Error: Credentials have not been saved prior. Please use Save-Credentials and try again" -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced object does not exist'
        Break
    }
    If ($Source)
    {
        Write-Host "This command utilizes a custom PSObject to offer a selection of services to connect to, and then calls the corresponding commands to connect to the requested service. `nFull details are unavailable through this help prompt."
    }
    Else
    {
        Write-Host "Connecting to multiple services using saved credentials"
        $Credential = Import-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
        Read-Host "Press enter to select the services to connect to" 
        $PowerShellOption = @([pscustomobject]@{Name="AzureAD"},[pscustomobject]@{Name="MSOL"},[pscustomobject]@{Name="SecurityCompliance"},[pscustomobject]@{Name="ExchangeOnline"},[pscustomobject]@{Name="Skype"},[pscustomobject]@{Name="SharePointPNP"},[pscustomobject]@{Name="SharepointAdmin"})
        $PowershellSelection = $PowerShellOption |out-gridview -PassThru -Title "Select Up to 5 Services to connect to" 
        If($PowershellSelection.Count -le 5)
        {
            If($PowershellSelection.Name.Contains('AzureAD'))
            {
                Connect-Azure
            }
            If($PowershellSelection.Name.Contains('MSOL'))
            {
                Connect-MSOL
            }
            If($PowershellSelection.Name.Contains('SecurityCompliance'))
            {
                Connect-Security
            }
            If($PowershellSelection.Name.Contains('ExchangeOnline'))
            {
                Connect-Exchange
            }
            If($PowershellSelection.Name.Contains('Skype'))
            {
                Connect-Skype
            }
            If($PowershellSelection.Name.Contains('SharePointPNP'))
            {
                Connect-SharePointPnP
            }
            If($PowershellSelection.Name.Contains('SharepointAdmin'))
            {
                Connect-SharePoint
            }
        }
        Else
        {
            Write-Host "Too many values selected please call Connect-Multi again" 
        }
    }
}

#This function is designed to save credentials in a secure manner on the local machine, anyone with access to account on the machine will be able to access the credentials, however they cannot be copied or reused on another machine. 
function Save-Credentials
{
    param([Switch]$Source)
    If (-not( Test-Path -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\'))
    {
        Write-Error -Message "Error: Path Does not exist, Please ensure the C:\Program Files\WindowsPowerShell\Configuration\Registration\ folder has been created')) " -Category ObjectNotFound -CategoryActivity 'File Access' -CategoryReason 'Referenced Folder does not exist'
        Break
    }
    If($Source)
    {
        Write-Host "Saves credentials to the PowerShell Registration folder using Export-CliXml Review Example 3:`nhttps://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-clixml`nNote Credential files are encrypted using the Windows Data Protection API and will only run on the machine from which they were saved Source Code: `n`$credential = Get-Credential`n`$credential | Export-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'"
    }
    Else
    {
        Write-Host "Saving Credentials"
        $credential = Get-Credential
        $credential | Export-CliXml -Path 'C:\Program Files\WindowsPowerShell\Configuration\Registration\credentials.xml'
    }
}

#This function provides the core setup and installation for all Office 365 Adminstration tools
#This does not include teams as the module currently provides no additional benefit over the GUI as of the time of writing. 
function Install-AdminTool
{
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    Read-Host "Press Enter to download the requisite modules for Office Administration"
    Install-Module -Name AzureAD -Force -ErrorAction SilentlyContinue
    Install-Module -Name MSOnline -Force
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
    Install-Module -Name SharePointPnPPowerShellOnline -Force
    Read-Host "Please Download and install the Skype Online Connector from Here: `nhttps://www.microsoft.com/en-us/download/details.aspx?id=39366`nPress Enter to continue"
    Import-Module -Name SkypeOnlineConnector -Force
    Read-Host "Please download the Microsoft Online Services Sign-In Assistant for IT Professionals RTW from here:`nhttps://www.microsoft.com/en-us/download/details.aspx?id=41950`nPress Enter to continue."
}

#This function updates any module downloaded from PSGallery
function Update-AdminTool
{
    Read-Host "Press Enter to update all modules for Office Administration"
    Update-Module -Name AzureAD -Force -ErrorAction SilentlyContinue
    Update-Module -Name MSOnline -Force
    Update-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
    Update-Module -Name SharePointPnPPowerShellOnline -Force
}

#This function provides a session management utility to reduce the chance of lockout from too many active sessions.
Function Logout-All
{
    param([Switch]$Source,[Switch]$Exit)
    If($Source)
    {
        Write-Host "Collects then removes any active PSSessions to prevent login lockout specify -Exit to exit PowerShell`nSource is: Get-PSSession |Remove-PSSession" 
    }
    If($Exit)
    {
        Get-PSSession |Remove-PSSession
        Exit
    }
    Else
    {
        Get-PSSession |Remove-PSSession
    }
}