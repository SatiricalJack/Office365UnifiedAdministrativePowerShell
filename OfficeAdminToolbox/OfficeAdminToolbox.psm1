<#
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

By installing, you have acknowledged the above disclaimer. Do not remove this section.
#>

function Clean-ProductKey
{
    param([Switch]$Status,[String]$Key)
    
    if($key)
    {
        If ($Key.Length -NE 5)
        {
            Write-Error -Message "Error, Key Length is too short, please ensure that all 5 characters have been entered" 
            Break
        }
        else
        {
            $Dir = Get-Location 
            Cd "C:\program files (x86)\Microsoft Office\Office16" -ErrorAction SilentlyContinue
            Cd "C:\program files\Microsoft office\office16" -ErrorAction SilentlyContinue
            cscript ospp.vbs /unpkey:$Key
            Cd $Dir.Path
        }
    }
    if($Status)
    {
        $Dir = Get-Location 
        Cd "C:\program files (x86)\Microsoft Office\Office16" -ErrorAction SilentlyContinue
        Cd "C:\program files\Microsoft office\office16" -ErrorAction SilentlyContinue
        #Allows for Automatic Selection of correct file path for default 32 or 64 bit installations.

        cscript ospp.vbs /dstatus |Out-File C:\ProductKeyStatus.txt
        Write-Host "Review file located at C:\ProductKeyStatus.txt"
        Cd $Dir.Path
    }
    
    Else
    {
        $Dir = Get-Location 
        Cd "C:\program files (x86)\Microsoft Office\Office16" -ErrorAction SilentlyContinue
        Cd "C:\program files\Microsoft office\office16" -ErrorAction SilentlyContinue
        #Allows for Automatic Selection of correct file path for default 32 or 64 bit installations.

        cscript ospp.vbs /dstatus |Out-File C:\ProductKeyStatus.txt
        Start-Sleep -s 5
        #Writes to a file to parse later
    
        #Confirms Operation to remove keys
        $input=$null
        $input = Read-Host "WARNING This will remove all installed Product Keys from Office Products. `nThis Operation Cannot be Undone Type y to confirm, anything else to quit"

        #Parses Extracts key information, and removes all keys 
        $Output = get-content C:\ProductKeyStatus.txt | Select-String -Pattern "Last 5 characters of installed product key: "
        if($input -eq "y")
        {
            $Output -replace "Last 5 characters of installed product key: ", ""|ForEach-Object {cscript ospp.vbs /unpkey:$_}
        }
        #Exit Case for cancellation
        Else
        {
            Write-Host "Your installed product keys have not been removed"
        }
        #Cleanup
        Remove-Item –path C:\ProductKeyStatus.txt
        Cd $Dir.Path
    }
}


function Get-GroupMembers
{
    Param([Switch]$AsFile, [Switch]$Source)
    If ($AsFile)
    {
        $GroupIdentity = Read-Host "Enter the name of the group to export to a CSV"
        get-distributiongroupmember -identity $GroupIdentity | foreach-object{New-Object psobject -property @{alias = $_.Alias;Identity = $_.identity;Guid = $_.ExchangeGuid; DisplayName = $_.DisplayName; EmailAddresses = $_.EmailAddresses}}| Export-Csv -Path C:\Mailboxes.csv -NoTypeInformation
        Write-Host "Please review file located at C:\Mailboxes.csv"
    }
    if ($Source)
    {
    Write-Host "This is used to gather important information regarding distribution group members for a more detailed investigation later `nSource Code:  get-distributiongroupmember -identity <GroupIdentity> | foreach-object{New-Object psobject -property @{alias = `$_.Alias;Identity = `$_.identity;Guid = `$_.ExchangeGuid; DisplayName = `$_.DisplayName; EmailAddresses = `$_.EmailAddresses}}| Export-Csv -Path C:\Mailboxes.csv -NoTypeInformation"
    }
    Else
    {
      $GroupIdentity = Read-Host "Enter the name of the group to export to a CSV"
      get-distributiongroupmember -identity $GroupIdentity | foreach-object{New-Object psobject -property @{alias = $_.Alias;Identity = $_.identity;Guid = $_.ExchangeGuid; DisplayName = $_.DisplayName; EmailAddresses = $_.EmailAddresses}}    
    }
}