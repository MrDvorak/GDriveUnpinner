# Get script startup time
$start = (Get-Date).Second

# Get Shell App
$app = New-Object -ComObject shell.application

function Get-GDrive
{
    return $app.Namespace("shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}").Items() | Where-Object { $_.Name -match '^Google Drive \([A-Z]:\)$' }
}

# Wait for GDrive to Quick access pipping
while (!$(Get-GDrive))
{
    # Does we wait too long?
    if ((Get-Date).Second - $start -gt 60*3)
    {
	Write-Verbose "GDrive wasn't found!"
        exit
    }

    Write-Verbose "Waiting for GDrive..."
    Start-Sleep -Seconds 3
}

# Remove GDrive from quick access:
if ($gdrive = Get-GDrive)
{
    $gdrive.InvokeVerb("unpinfromhome")
}
