param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$word = $null
$document = $null

try {
    try {
        $word = New-Object -ComObject Word.Application
    }
    catch {
        exit 2
    }

    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($InputPath, $false, $true, $false)
    $document.Repaginate()
    $document.Fields.Update()
    $document.ExportAsFixedFormat(
        $OutputPath,
        17,
        $false,
        0,
        0,
        1,
        1,
        $false,
        $false,
        $false,
        $false,
        $false,
        $false
    )
}
catch {
    exit 1
}
finally {
    if ($null -ne $document) {
        try {
            $document.Close(0)
        }
        catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
        catch {}
    }
    if ($null -ne $word) {
        try {
            $word.Quit(0)
        }
        catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        }
        catch {}
    }
}
