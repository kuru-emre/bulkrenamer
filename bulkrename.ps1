$Shell = New-Object -ComObject shell.application

# Iterate through each file
Get-ChildItem -Recurse -file *.wav -ErrorAction Stop | ForEach{

    # Skip the WhatsApp images (Remove this block if you don't want to skip)
    if ($_.Name.StartsWith("WhatsApp Image")) {
        continue
    }

    $Folder = $Shell.NameSpace($_.DirectoryName)
    $File = $Folder.ParseName($_.Name)

    # Find the available property
    $Property = $Folder.GetDetailsOf($File,12)
    if (-not $Property) {
        $Property = $Folder.GetDetailsOf($File,3)
        if (-not $Property) {
            $Property = $Folder.GetDetailsOf($File,4)
        }
    }

    # Get date in the required format as a string
    $RawDate = ($Property -Replace "[^\w /:]")
    $DateTime = [DateTime]::Parse($RawDate)
    $DateTaken = $DateTime.ToString("yyyy-MM-dd HH.mm")

    # Find the available file name and duplicate names
    $Iterator = 1
    $FileName = $DateTaken
    $Path = $_.DirectoryName + "\" + $FileName + $_.Extension
    while (Test-Path -Path $Path -PathType Leaf) {
        $FileName = $DateTaken + " (" + $Iterator + ")"
        $Path = $_.DirectoryName + "\" + $FileName + $_.Extension
        $Iterator++
    }

    # Rename file
    Write-Output $_.Name"=>"$FileName
    Rename-Item $_.FullName ($FileName + $_.Extension)
}