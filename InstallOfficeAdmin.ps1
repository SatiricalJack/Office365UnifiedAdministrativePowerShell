﻿Read-host "THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.`nIN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE`nSOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.`n`nBy continuing to install, you acknowledge the above disclaimer. A copy of this will also be made available within the installed software.`n`nPress Enter to continue."
copy-item -path ./OfficeAdmin -destination ($ENV:USERPROFILE+'\Documents\WindowsPowerShell\Modules') -recurse -Force
copy-item -path ./OfficeAdminHelp -destination ($ENV:USERPROFILE+'\Documents\WindowsPowerShell\Modules') -recurse -Force
copy-item -path ./OfficeAdminToolbox -destination ($ENV:USERPROFILE+'\Documents\WindowsPowerShell\Modules') -recurse -Force