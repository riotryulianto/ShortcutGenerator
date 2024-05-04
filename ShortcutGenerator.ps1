# Directory where shortcuts will be created
$shortcutFolder = "C:\Users\$ENV:UserName\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Portable"

# Directory containing .exe files (change this to your desired directory)
$sourceFolder = $PSScriptRoot

# Directories to exclude (change or add more as needed)
$excludedFolders = @(
	"App",
	"Data",
	"Other",
	"System",
	"python",
	"CdiResource",
	"CdmResource",
	"jre",
	"plugins",
	"Source",
	"SP_Download_tool",
	"data",
	"tools",
	"altexe",
    "NirSoft",
    "SysinternalsSuite"
)

# Create the shortcut folder if it doesn't exist
New-Item -ItemType Directory -Force -Path $shortcutFolder

# Find all .exe files in the source folder and its subdirectories, excluding specified folders
$exeFiles = Get-ChildItem -Path $sourceFolder -Recurse -Filter *.exe | Where-Object { $_.FullName -notmatch ($excludedFolders -join '|') }

foreach ($exeFile in $exeFiles) {
    # Get the "File description" property from the .exe file
    $fileDescription = (Get-ItemProperty $exeFile.FullName).VersionInfo.FileDescription

    if (-not [string]::IsNullOrWhiteSpace($fileDescription)) {
        # Check if the length of the file description is greater than 50 characters
        if ($fileDescription.Length -gt 50) {
            # Use the folder name as the base shortcut name
            $baseShortcutName = Join-Path -Path $shortcutFolder -ChildPath "$($exeFile.Directory.Name)"
        } else {
            # Use the file description as the base shortcut name
            $baseShortcutName = Join-Path -Path $shortcutFolder -ChildPath "$fileDescription"
        }
    } else {
        # Use the file name as the base shortcut name if file description is empty
        $baseShortcutName = Join-Path -Path $shortcutFolder -ChildPath "$($exeFile.BaseName)"
    }

    # Check if a shortcut with the base name already exists
    $shortcutNumber = 1
    $shortcutName = "$baseShortcutName.lnk"
    while (Test-Path $shortcutName) {
        $shortcutNumber++
        $shortcutName = "$baseShortcutName ($shortcutNumber).lnk"
    }

    # Create the shortcut
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut($shortcutName)
    $shortcut.TargetPath = $exeFile.FullName
    $shortcut.Save()
    Write-Output "Shortcut created for $($exeFile.FullName)"
}
