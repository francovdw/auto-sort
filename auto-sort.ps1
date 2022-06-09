# Force User Access Control for script
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Archive folder location
$path = "$env:USERPROFILE\Documents\_Archive"

# Start Transcript
Write-Host ""
Start-Transcript -path "$path\trans.txt" -Append
Write-Host ""
Write-Host ""

# Create folders based on the extension list
$confirmfolder = Test-Path -PathType Container "$path\_Other"
    if ($confirmfolder -ne $true){New-Item "$path\_Other" -ItemType Directory -Force -EA SilentlyContinue}

# File name extension list (File types that will be moved to the archived folder)
$Extension_List = "*.ps1","*.txt","*.pdf","*.docx","*.doc","*.xlsx","*.xls","*.png","*.jpg","*.csv","*.exe","*.msi","*.zip","*.rar","*.tar","*.log","*.cfg","*.iso","*.rdp","*.mp4","*.mp3","*.wav","*.msg","*.cer","*.deb","*.bat","*.pfx","*.img","*.ini","*.xlsm","*.pptx","*.msu","*.tbxml","*.pcapn","*.drawio"

# File name extension Exclude list (Files to be excluded from being moved)
# PS: Add Name & File name extension
$Exclude_List = "auto-sort.lnk","auto-sort.ps1","AC.lnk","New Tab.lnk","Chrome.lnk","seam.lnk"

# Script
foreach($extension in $Extension_List)
    {
        # Create folder name from extension
        $folder = $extension.TrimStart("*.")
        $folder = $folder.ToUpper()

        # Create folder for the file type
        $ConfirmFolder = Test-Path -PathType Container "$path\$folder" -Verbose
            if ($confirmfolder -ne $true){New-Item "$path\$folder" -ItemType Directory -Force -EA SilentlyContinue -Verbose}

        # Get files from Downloads & Desktop
        $Downloads = "$env:USERPROFILE\Downloads\$extension"
        $Desktop = "$env:USERPROFILE\Desktop\$extension"

        # Move file to folder
        Move-Item -Path $Downloads -Destination "$path\$folder" -Exclude $Exclude_List -Force -Verbose -EA SilentlyContinue
        Move-Item -Path $Desktop -Destination "$path\$folder" -Exclude $Exclude_List -Force -Verbose -EA SilentlyContinue

            }

# Move all unspecified files & folders into one folder
Move-Item -Path "$env:USERPROFILE\Downloads\*" -Destination "$Path\_Other" -Exclude $Exclude_List -Force -Verbose -EA SilentlyContinue
Move-Item -Path "$env:USERPROFILE\Desktop\*" -Destination "$Path\_Other" -Exclude $Exclude_List -Force -Verbose -EA SilentlyContinue

# Stop Transcript
Write-Host ""
Write-Host ""
Stop-Transcript
Write-Host ""
Write-Host ""
Write-Host ""
pause
exit
